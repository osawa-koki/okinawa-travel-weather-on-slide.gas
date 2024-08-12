function insertOrUpdateTextBox() {
  // プロパティストアからプレゼンテーションID、スライドID、およびオブジェクトIDを取得
  var properties = PropertiesService.getScriptProperties();
  var presentationId = properties.getProperty('PRESENTATION_ID');
  var slideId = properties.getProperty('SLIDE_ID');
  var objectId = properties.getProperty('OBJECT_ID'); // オブジェクトID

  if (!presentationId || !slideId) {
    Logger.log('Presentation ID or Slide ID is not set.');
    return;
  }

  // プレゼンテーションオブジェクトを取得
  var presentation = SlidesApp.openById(presentationId);
  var slide = presentation.getSlides().find(slide => slide.getObjectId() === slideId);

  // オブジェクトIDがプロパティにあり、スライド内にも存在する場合、更新
  var shape = objectId ? slide.getShapes().find(shape => shape.getObjectId() === objectId) : null;

  if (shape) {
    // 既存のオブジェクトを更新
    shape.getText().setText('Updated text in Google Slides!');
    Logger.log('Text box with ID: ' + objectId + ' has been updated.');
  } else {
    // 新規にテキストボックスを作成し、プロパティに保存
    shape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 100, 100, 300, 100);
    shape.getText().setText('New text in Google Slides!');
    var newObjectId = shape.getObjectId();
    properties.setProperty('OBJECT_ID', newObjectId); // プロパティにオブジェクトIDを保存
    Logger.log('New text box created with ID: ' + newObjectId);
  }
}
