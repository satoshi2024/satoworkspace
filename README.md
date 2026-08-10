const viewer = document.querySelector("image-viewer");
const bg = viewer.shadowRoot.querySelector("#background-layer");
const ui = viewer.shadowRoot.querySelector("#ui-layer");

console.log("background attribute:",
    bg.getAttribute("width"),
    bg.getAttribute("height"));

console.log("background real:",
    bg.width,
    bg.height);

console.log("background display:",
    bg.getBoundingClientRect().width,
    bg.getBoundingClientRect().height);

console.log("ui attribute:",
    ui.getAttribute("width"),
    ui.getAttribute("height"));

console.log("ui real:",
    ui.width,
    ui.height);

console.log("ui display:",
    ui.getBoundingClientRect().width,
    ui.getBoundingClientRect().height);

console.log("devicePixelRatio:",
    window.devicePixelRatio);