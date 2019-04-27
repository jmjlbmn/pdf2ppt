# pdf2ppt
pdf转换ppt,简单实现(舍弃原JavaFX实现的GUI版本)
- 基于pdfbox和poi实现
- 生成的ppt是纯图片的
- 两种方式
  - 将pdf渲染为图片,然后转换ppt(兼容更好)
  - 提取pdf的原图,转换ppt(质量更好,仅适用于纯图片/扫描pdf)
- 提供一个pdf导出所有图片的方法