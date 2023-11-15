import pptxgen from "npm:pptxgenjs";
import sizeOf from 'npm:image-size';

createPPT('1212', ['pics/1.jpg', 'pics/2.jpg'])

function createPPT(fileName, pics) {
  let pptx = new pptxgen();
  const width = 11.69;
  const height = 8.27;
  // 初始化纸张大小
  pptx.defineLayout({
    name: 'A4',
    width: width,
    height: height
  });
  pptx.layout = 'A4';

  for (let index = 0; index < pics.length; index++) {
    const pic = pics[index];
    pptx.addSlide().addImage(genImagesParams(pic, width, height));
  }

  pptx.writeFile({
    fileName: `${fileName}.pptx`
  });
}

function genImagesParams(picPath, pageWidth, pageHeight) {
  const metadata = sizeOf(picPath)
  var params = {
    path: picPath,
    x: 0,
    y: 0,
    w: 0,
    h: 0
  }
  if (metadata.width < metadata.height) {
    // 只保证高度撑满页面即可，宽度随意
    params.h = pageHeight
    params.w = parseFloat((pageHeight * (metadata.width / metadata.height)))
    params.x = parseFloat(((pageWidth  - params.w) / 2))
  } else {
    // 保证宽度，高度居中(可裁剪)
    params.w = pageWidth
    params.h = parseFloat((pageWidth * (metadata.height / metadata.width)))
    params.y = parseFloat(((pageHeight  - params.h) / 2))
  }
  return params;
}
