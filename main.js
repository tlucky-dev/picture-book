import pptxgen from "npm:pptxgenjs";
import sizeOf from 'npm:image-size';

createPPT('1212', ['pics/1.jpg', 'pics/2.jpg'])

function createPPT(fileName, pics) {
  let pptx = new pptxgen()
  // 初始化纸张大小
  pptx.defineLayout({ name:'A4', width:11.69, height:8.27 });
  pptx.layout = 'A4';

  for (let index = 0; index < pics.length; index++) {
    const pic = pics[index];
    pptx.addSlide().addImage(genImagesParams(pic));
  }

  pptx.writeFile({
    fileName: `${fileName}.pptx`
  });
}

function genImagesParams(picPath) {
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
    params.h = 8.27
    params.w = parseFloat((8.27 * (metadata.width / metadata.height)))
    params.x = parseFloat(((11.69  - params.w) / 2))
  } else {
    // 保证宽度，高度居中(可裁剪)
    params.w = 11.69
    params.h = parseFloat((11.69 * (metadata.height / metadata.width)))
    params.y = parseFloat(((8.27  - params.h) / 2))
  }
  return params;
}