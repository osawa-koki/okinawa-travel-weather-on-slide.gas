// @ts-nocheck

function insertOrUpdateTextBox() {
  const properties = PropertiesService.getScriptProperties()

  const presentationId = properties.getProperty('PRESENTATION_ID')
  const slideDateMapper = JSON.parse(properties.getProperty('SLIDE_DATE_MAPPER'))

  if (presentationId == null) {
    Logger.log('Presentation ID is not set.')
    return
  }

  const presentation = SlidesApp.openById(presentationId)

  for (const [i, { slideId, date }] of Object.entries(slideDateMapper)) {
    const objectIdKey = `OBJECT_ID_${i}`
    const objectId = properties.getProperty(objectIdKey)

    if (slideId == null || date == null) {
      Logger.log('Slide ID or date is not set.')
      return
    }

    Logger.log('Slide ID: ' + slideId + ', Date: ' + date)

    const slide = presentation.getSlides().find(slide => slide.getObjectId() === slideId)

    if (slide == null) {
      Logger.log('Slide with ID: ' + slideId + ' not found.')
      return
    }

    let shape = objectId ? slide.getShapes().find(shape => shape.getObjectId() === objectId) : null

    const text = '🌞'
    const {x, y, width, height} = {x: 650, y: 50, width: 50, height: 50}

    if (shape != null) {
      shape.getText().setText(text)
      shape.setLeft(x)
      shape.setTop(y)
      shape.setWidth(width)
      shape.setHeight(height)
      Logger.log('Text box with ID: ' + objectId + ' has been updated.')
    } else {
      shape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, y, width, height)
      shape.getText().setText(text)
      const newObjectId = shape.getObjectId()
      properties.setProperty(objectIdKey, newObjectId)
      Logger.log('New text box created with ID: ' + newObjectId)
    }
  }
}
