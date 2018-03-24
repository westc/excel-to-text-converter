[
  16, 32, 48, 64, 128, 192, 256, 512, 1024
].forEach(function(size) {
  var fontSize = size < 64 size / 5 : 0, height = size, width = size;
  var data = `
    <svg xmlns="http://www.w3.org/2000/svg" height="${height}" width="${width}">
      <foreignObject width="100%" height="100%">
        <div xmlns="http://www.w3.org/1999/xhtml">
          <div style="background-image: radial-gradient(#07F, transparent, #07F, black 75%); display: table; width: 5em; height: 5em; font-size: ${fontSize}px; border-radius: 0 5em; box-shadow: inset 0 0 0.125em 1px #000, inset 0 0 0.25em 5px #FFF;">
            <div style="display: table-cell; vertical-align: middle; text-align: center;">
              <div style="font-family: 'Trebuchet MS'; letter-spacing: 0.2em; font-weight: bold; font-size: 0.8em; color: #FFF; text-shadow: 0 0 0.1em #000, 0 0 0.1em #000, 0 0 0.1em #000, 0 0 0.1em #000; display: inline-block; position: relative;">
                <div style="content: ''; position: absolute; top: 0; left: 0; right: 0; bottom: 0; border-radius: 100%; box-shadow: -1.5em 0 1em -1em #000; z-index: -1;"></div>
                Excel<br/>&#8681;<br/>Text
                <div style="content: ''; position: absolute; top: 0; left: 0; right: 0; bottom: 0; border-radius: 100%; box-shadow: 1.5em 0 1em -1em #000; z-index: -1;"></div>
              </div>
            </div>
          </div>
        </div>
      </foreignObject>
    </svg>`;

  // Define an image
  var img = new Image();

  // Once the image loads draw it onto a new canvas
  img.onload = function() {
    console.log(25);
    // Create a canvas
    var canvas = document.createElement('canvas');
    canvas.style.height = 0;
    canvas.height = height;
    canvas.width = width;

    // Draw the image onto the canvas
    canvas.getContext('2d').drawImage(img, 0, 0);

    var imgToAppend = new Image();
    imgToAppend.style.display = 'block';
    imgToAppend.src = canvas.toDataURL();
    document.body.appendChild(imgToAppend);
  };

  // Set the source of the image to the SVG data
  img.src = 'data:image/svg+xml;base64,' + window.btoa(data);
});
