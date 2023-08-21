/* global document, Office,PowerPoint,Buffer */
let isDragging = false;
let startAngle = 0;
Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("apply-settings").addEventListener("click", insertWatermark);
    const dialContainer = document.querySelector(".dial-container");
    const rotationDegreeInput = document.getElementById("rotation-degree");
    const dialMarker = document.querySelector(".dial-marker");
    rotationDegreeInput.addEventListener("input", updateRotation);
    dialMarker.addEventListener("drag", updateDialRotation);
    dialContainer.addEventListener("mousedown", startDrag);
    dialContainer.addEventListener("touchstart", startDrag);

    // Attach input event handler to Font Size slider

    const fontSizeSlider = document.getElementById("font-size");
    fontSizeSlider.addEventListener("input", updateFontSizeOutput);
  }
});
function updateRotation() {
  const rotationDegreeInput = document.getElementById("rotation-degree");
  const dialMarker = document.querySelector(".dial-marker");

  const degrees = rotationDegreeInput.value;
  dialMarker.style.transform = `translateX(-50%) translateY(-100%) rotate(${degrees}deg)`;
}

function updateDialRotation(event) {
  const rotationDegreeInput = document.getElementById("rotation-degree");
  const dialMarker = document.querySelector(".dial-marker");

  const rect = event.target.getBoundingClientRect();
  const centerX = rect.left + rect.width / 2;
  const centerY = rect.top + rect.height / 2;
  const x = event.clientX - centerX;
  const y = centerY - event.clientY;
  const degrees = Math.atan2(y, x) * (180 / Math.PI) + 90;
  rotationDegreeInput.value = degrees;
  dialMarker.style.transform = `translateX(-50%) translateY(-100%) rotate(${degrees}deg)`;
}
// Update Font Size output element with slider value
function updateFontSizeOutput() {
  const fontSizeSlider = document.getElementById("font-size");
  const fontSizeOutput = document.getElementById("font-size-output");
  fontSizeOutput.textContent = fontSizeSlider.value;
}

async function insertWatermark() {
  //await deleteWatermark();
  await PowerPoint.run(async (context) => {
    const watermarkText = document.getElementById("watermark-text").value;
    const font = document.getElementById("font-select").value;
    const fontSize = parseInt(document.getElementById("font-size").value, 10);
    const fontColor = document.getElementById("font-color").value;
    const rotationAngle = document.getElementById("rotation-degree").value;

    const svgWidth = 1600;
    const svgHeight = 900;

    const centerX = svgWidth / 2;
    const centerY = svgHeight / 2;

    /*const svgMarkup = `
  <svg width="${svgWidth}" height="${svgHeight}">
    <text x="${centerX}" y="${centerY}" font-size="${fontSize}" fill="${fontColor}" font-family="${font}"
      text-anchor="middle" alignment-baseline="middle"
      transform="rotate(315 ${centerX} ${centerY})" opacity="0.5">${watermarkText}</text>
  </svg>
`;*/
    const svgMarkup = convertTextToSvgMarkup(
      watermarkText,
      svgWidth,
      svgHeight,
      rotationAngle,
      centerX,
      centerY,
      fontSize,
      fontColor,
      font
    );
    const base64EncodedSVG = btoa(unescape(encodeURIComponent(svgMarkup)));

    // Convert SVG markup to a data URL
    //var dataUrl = "data:image/svg+xml;base64," + customBtoa(unescape(encodeURIComponent(svgMarkup)));
    //console.log(base64EncodedSVG);
    // Insert the image from the base64-encoded SVG
    // Iterate through each slide
    const slides = context.presentation.slides;
    const slideCount = slides.getCount();
    await context.sync();
    for (let i = 0; i < slideCount.value; i++) {
      //const slide = slides.getItemAt(i);
      //context.presentation.goToSlide(i);
      await insertImageFromBase64String(base64EncodedSVG);
      goToNextSlide();
    }
    goToFirstSlide();

    await context.sync();

    slides.load("items/$none");

    await context.sync();
    slides.items.forEach(async (slide, index) => {
      const shapes = slide.shapes;
      // Load all the shapes in the collection without loading their properties.
      shapes.load("items");

      await context.sync();
      const shape = shapes.items[shapes.items.length - 1];
      shape.name = `watermark_${index}`;
    });
    await context.sync();
  });
}
function convertTextToSvgMarkup(
  watermarkText,
  svgWidth,
  svgHeight,
  rotationAngle,
  centerX,
  centerY,
  fontSize,
  fontColor,
  font
) {
  // Split the watermarkText by newline characters
  const lines = watermarkText.split("\n");

  // Create an array of <tspan> elements for each line
  const tspans = lines.map((line, index) => {
    const y = centerY + (index - Math.floor(lines.length / 2)) * fontSize; // Adjust y position for each line
    return `<tspan x="${centerX}" y="${y}" text-anchor="middle" alignment-baseline="middle">${line}</tspan>`;
  });

  // Combine the tspans into the SVG markup
  const svgMarkup = `
    <svg width="${svgWidth}" height="${svgHeight}">
      <text font-size="${fontSize}" fill="${fontColor}" font-family="${font}"
        text-anchor="middle" alignment-baseline="middle"
        transform="rotate(${rotationAngle} ${centerX} ${centerY})" opacity="0.5">
        ${tspans.join("\n")}
      </text>
    </svg>
  `;

  return svgMarkup;
}
function goToNextSlide() {
  Office.context.document.goToByIdAsync(Office.Index.Next, Office.GoToType.Index, function (asyncResult) {
    if (asyncResult.status == "failed") {
    }
  });
}
function goToFirstSlide() {
  Office.context.document.goToByIdAsync(Office.Index.First, Office.GoToType.Index, function (asyncResult) {
    if (asyncResult.status == "failed") {
    }
  });
}
async function deleteWatermark() {
  await PowerPoint.run(async (context) => {
    // Delete all shapes from the first slide.
    const sheet = context.presentation.slides.getItemAt(0);
    const shapes = sheet.shapes;

    // Load all the shapes in the collection without loading their properties.
    shapes.load("items/$none");
    await context.sync();

    shapes.items.forEach(function (shape) {
      shape.delete();
    });
    await context.sync();
  });
}
async function insertImageFromBase64String(image) {
  return new Promise((resolve) => {
    // Call Office.js to insert the image into the document.
    Office.context.document.setSelectedDataAsync(
      image,
      {
        coercionType: Office.CoercionType.Image,
      },
      function (asyncResult) {
        const insertedImage = asyncResult.value;
        resolve(insertedImage);
      }
    );
  });
}
function startDrag(event) {
  event.preventDefault();
  isDragging = true;

  startAngle = getCurrentAngle(event);

  document.addEventListener("mousemove", drag);
  document.addEventListener("touchmove", drag);

  document.addEventListener("mouseup", stopDrag);
  document.addEventListener("touchend", stopDrag);
}

function drag(event) {
  const rotationDegreeInput = document.getElementById("rotation-degree");
  // const rotationDegreeOutput = document.getElementById("rotation-degree-output");
  const dialMarker = document.querySelector(".dial-marker");

  if (!isDragging) return;

  const currentAngle = getCurrentAngle(event);
  const rotationAngle = currentAngle - startAngle;

  dialMarker.style.transform = `translateX(-50%) translateY(-100%) rotate(${rotationAngle}deg)`;
  rotationDegreeInput.value = Math.round(rotationAngle);
  // rotationDegreeOutput.textContent = Math.round(rotationAngle);
}

function stopDrag() {
  isDragging = false;

  document.removeEventListener("mousemove", drag);
  document.removeEventListener("touchmove", drag);

  document.removeEventListener("mouseup", stopDrag);
  document.removeEventListener("touchend", stopDrag);
}

function getCurrentAngle(event) {
  const dialContainer = document.querySelector(".dial-container");

  const dialRect = dialContainer.getBoundingClientRect();
  const centerX = dialRect.left + dialRect.width / 2;
  const centerY = dialRect.top + dialRect.height / 2;

  const mouseX = event.clientX || event.touches[0].clientX;
  const mouseY = event.clientY || event.touches[0].clientY;

  const angle = Math.atan2(mouseY - centerY, mouseX - centerX);
  const degrees = (angle * 180) / Math.PI;

  return degrees;
}
