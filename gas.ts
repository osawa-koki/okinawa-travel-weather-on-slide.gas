// @ts-nocheck

async function getWeatherData({ openWeatherMapApi }) {
  const response = UrlFetchApp.fetch(`https://api.openweathermap.org/data/2.5/forecast?q=okinawa&appid=${openWeatherMapApi}`)
  const json = JSON.parse(response.getContentText())
  return json
}

function weatherToEmoji({ weather }) {
  switch (weather) {
    case 'Clear':
      return '🌞'
    case 'Clouds':
      return '🌥️'
    case 'Rain':
      return '🌧'
    case 'Snow':
      return '⛄️'
    case 'Thunderstorm':
      return '⚡️'
    default:
      return '?'
  }
}

function getWeather({ weatherData, date }) {
  const { list } = weatherData
  const items = list.map((item) => {
    const givenDate = new Date(date)
    const _date = new Date(item.dt_txt)
    if (givenDate.getDate() !== _date.getDate()) {
      return null
    }
    return weatherToEmoji({ weather: item.weather[0].main })
  }).filter(item => item != null)
  if (items.length === 0) return null
  return items
}

async function main() {
  const properties = PropertiesService.getScriptProperties()

  const presentationId = properties.getProperty('PRESENTATION_ID')
  const slideDateMapper = JSON.parse(properties.getProperty('SLIDE_DATE_MAPPER'))

  const openWeatherMapApi = properties.getProperty('OPEN_WEATHER_MAP_API')

  if (presentationId == null) {
    Logger.log('Presentation ID is not set.')
    return
  }

  const presentation = SlidesApp.openById(presentationId)

  const weatherData = await getWeatherData({ openWeatherMapApi })
  Logger.log('Weather data: ' + JSON.stringify(weatherData))

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

    const weather = getWeather({ weatherData, date })

    const text = weather != null
      ? `${weather.join(' ')}\n(${date})`
      : `No data\n(${date})`
    const { x, y, width, height } = { x: 450, y: 30, width: 250, height: 50 }

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
