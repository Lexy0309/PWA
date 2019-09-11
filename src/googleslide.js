export function createPresentation(title, callback) {
  window.gapi.client.slides.presentations
    .create({
      title: title
    })
    .then(response => {
      console.log(
        `Created presentation with ID: ${response.result.presentationId}`
      );
      callback(response);
    });
}

export function copyPresentation(presentationId, copyTitle, callback) {
  var request = {
    name: copyTitle
  };
  window.gapi.client.drive.files
    .copy({
      fileId: presentationId,
      resource: request
    })
    .then(driveResponse => {
      var presentationCopyId = driveResponse.result.id;
      callback(presentationCopyId);
    });
}

export function createSlide(
  presentationId,
  insertionIndex,
  predefinedLayout,
  pageId,
  callback
) {
  var requests = [
    {
      createSlide: {
        objectId: pageId,
        insertionIndex: insertionIndex,
        slideLayoutReference: {
          predefinedLayout: predefinedLayout
        }
      }
    }
  ];

  window.gapi.client.slides.presentations
    .batchUpdate({
      presentationId: presentationId,
      requests: requests
    })
    .then(createSlideResponse => {
      console.log(
        `Created slide with ID: ${
          createSlideResponse.result.replies[0].createSlide.objectId
        }`
      );
      callback(createSlideResponse);
    });
}

export function appendSlide(pid, layout, id) {
  let reqs = [
    {
      createSlide: {
        objectId: id,
        slideLayoutReference: {
          predefinedLayout: layout
        }
      }
    }
  ];

  return reqs;
}

export function createTextboxWithText(
  presentationId,
  pageId,
  id,
  txt,
  objStyle
) {
  let { size, offset, unit, bold, italic, fontSize } = objStyle;
  let elementId = id;
  let width = { magnitude: size.width, unit: unit };
  let height = { magnitude: size.height, unit: unit };
  let requests = [
    {
      createShape: {
        objectId: elementId,
        shapeType: "TEXT_BOX",
        elementProperties: {
          pageObjectId: pageId,
          size: { width, height },
          transform: {
            scaleX: 1,
            scaleY: 1,
            translateX: offset.x,
            translateY: offset.y,
            unit: unit
          }
        }
      }
    },
    {
      insertText: {
        objectId: elementId,
        insertionIndex: 0,
        text: txt
      }
    },
    {
      updateTextStyle: {
        objectId: elementId,
        textRange: { type: "ALL" },
        style: {
          fontFamily: "Times New Roman",
          bold: bold,
          italic: italic,
          fontSize: { magnitude: fontSize, unit: "PT" }
        },
        fields: "fontFamily,fontSize,bold,italic"
      }
    }
  ];
  return requests;
}

export function createImage(
  presentationId,
  imageId,
  pageId,
  imageFilePath,
  size,
  offset,
  UNIT = "PT"
) {
  var DISTANCE_UNIT = UNIT;
  var imageUrl = imageFilePath;
  let width = { magnitude: size.width, unit: DISTANCE_UNIT };
  let height = { magnitude: size.height, unit: DISTANCE_UNIT };
  var requests = [
    {
      createImage: {
        objectId: imageId,
        url: imageUrl,
        elementProperties: {
          pageObjectId: pageId,
          size: { width, height },
          transform: {
            scaleX: 1,
            scaleY: 1,
            translateX: offset.x,
            translateY: offset.y,
            unit: DISTANCE_UNIT
          }
        }
      }
    }
  ];
  return requests;
}

export function createVideo(
  presentationId,
  videoId,
  pageId,
  videoFilePath,
  size,
  offset,
  UNIT = "PT"
) {
  var DISTANCE_UNIT = UNIT;
  var videoUrl = videoFilePath;
  let width = { magnitude: size.width, unit: DISTANCE_UNIT };
  let height = { magnitude: size.height, unit: DISTANCE_UNIT };
  var requests = [
    {
      createVideo: {
        objectId: videoId,
        elementProperties: {
          pageObjectId: pageId,
          size: { width, height },
          transform: {
            scaleX: 1,
            scaleY: 1,
            translateX: offset.x,
            translateY: offset.y,
            unit: DISTANCE_UNIT
          }
        },
        source: "YOUTUBE",
        id: "7U3axjORYZ0"
      }

      // objectId: videoId,
      // url: videoUrl,
      // elementProperties: {
      //   pageObjectId: pageId,
      //   size: { width, height },
      //   transform: {
      //     scaleX: 1,
      //     scaleY: 1,
      //     translateX: offset.x,
      //     translateY: offset.y,
      //     unit: DISTANCE_UNIT
      //   }
      // }
    }
  ];
  return requests;
}

export function textMerging(
  templatePresentationId,
  dataSpreadsheetId,
  callback
) {
  var responses = [];
  var dataRangeNotation = "Customers!A2:M6";
  window.gapi.client.sheets.spreadsheets.values
    .get({
      spreadsheetId: dataSpreadsheetId,
      range: dataRangeNotation
    })
    .then(sheetsResponse => {
      var values = sheetsResponse.result.values;

      for (var i = 0; i < values.length; ++i) {
        var row = values[i];
        var customerName = row[2];
        var caseDescription = row[5];
        var totalPortfolio = row[11];
        var copyTitle = customerName + " presentation";
        var request = {
          name: copyTitle
        };
        window.gapi.client.drive.files
          .copy({
            fileId: templatePresentationId,
            requests: request
          })
          .then(driveResponse => {
            var presentationCopyId = driveResponse.result.id;

            var requests = [
              {
                replaceAllText: {
                  containsText: {
                    text: "{{customer-name}}",
                    matchCase: true
                  },
                  replaceText: customerName
                }
              },
              {
                replaceAllText: {
                  containsText: {
                    text: "{{case-description}}",
                    matchCase: true
                  },
                  replaceText: caseDescription
                }
              },
              {
                replaceAllText: {
                  containsText: {
                    text: "{{total-portfolio}}",
                    matchCase: true
                  },
                  replaceText: totalPortfolio
                }
              }
            ];

            window.gapi.client.slides.presentations
              .batchUpdate({
                presentationId: presentationCopyId,
                requests: requests
              })
              .then(batchUpdateResponse => {
                var result = batchUpdateResponse.result;
                responses.push(result.replies);
                var numReplacements = 0;
                for (var i = 0; i < result.replies.length; ++i) {
                  numReplacements +=
                    result.replies[i].replaceAllText.occurrencesChanged;
                }
                console.log(
                  `Created presentation for ${customerName} with ID: ${presentationCopyId}`
                );
                console.log(`Replaced ${numReplacements} text instances`);
                if (responses.length === values.length) {
                  callback(responses);
                }
              });
          });
      }
    });
}
export function imageMerging(
  templatePresentationId,
  imageUrl,
  customerName,
  callback
) {
  var logoUrl = imageUrl;
  var customerGraphicUrl = imageUrl;

  var copyTitle = customerName + " presentation";
  window.gapi.client.drive.files
    .copy({
      fileId: templatePresentationId,
      resource: {
        name: copyTitle
      }
    })
    .then(driveResponse => {
      var presentationCopyId = driveResponse.result.id;

      var requests = [
        {
          replaceAllShapesWithImage: {
            imageUrl: logoUrl,
            replaceMethod: "CENTER_INSIDE",
            containsText: {
              text: "{{company-logo}}",
              matchCase: true
            }
          }
        },
        {
          replaceAllShapesWithImage: {
            imageUrl: customerGraphicUrl,
            replaceMethod: "CENTER_INSIDE",
            containsText: {
              text: "{{customer-graphic}}",
              matchCase: true
            }
          }
        }
      ];

      window.gapi.client.slides.presentations
        .batchUpdate({
          presentationId: presentationCopyId,
          requests: requests
        })
        .then(batchUpdateResponse => {
          var numReplacements = 0;
          for (var i = 0; i < batchUpdateResponse.result.replies.length; ++i) {
            numReplacements +=
              batchUpdateResponse.result.replies[i].replaceAllShapesWithImage
                .occurrencesChanged;
          }
          console.log(
            `Created merged presentation with ID: ${presentationCopyId}`
          );
          console.log(`Replaced ${numReplacements} shapes with images.`);
          callback(batchUpdateResponse.result);
        });
    });
}

export function simpleTextReplace(
  presentationId,
  shapeId,
  replacementText,
  callback
) {
  var requests = [
    {
      deleteText: {
        objectId: shapeId,
        textRange: {
          type: "ALL"
        }
      }
    },
    {
      insertText: {
        objectId: shapeId,
        insertionIndex: 0,
        text: replacementText
      }
    }
  ];

  window.gapi.client.slides.presentations
    .batchUpdate({
      presentationId: presentationId,
      requests: requests
    })
    .then(batchUpdateResponse => {
      console.log(`Replaced text in shape with ID: ${shapeId}`);
      callback(batchUpdateResponse.result);
    });
}

export function emphasizeText(pid, shapeId) {
  let reqs = [
    {
      updateTextStyle: {
        objectId: shapeId,
        textRange: {
          type: "ALL"
        },
        style: {
          bold: true,
          fontFamily: "Times New Roman",
          fontSize: {
            magnitude: 24,
            unit: "PT"
          }
        },
        fields: "bold,fontFamily,fontSize"
      }
    }
  ];

  return reqs;
}

export function textStyleUpdate(presentationId, shapeId, callback) {
  var requests = [
    {
      updateTextStyle: {
        objectId: shapeId,
        textRange: {
          type: "FIXED_RANGE",
          startIndex: 0,
          endIndex: 5
        },
        style: {
          bold: true,
          italic: true
        },
        fields: "bold,italic"
      }
    },
    {
      updateTextStyle: {
        objectId: shapeId,
        textRange: {
          type: "FIXED_RANGE",
          startIndex: 5,
          endIndex: 10
        },
        style: {
          fontFamily: "Times New Roman",
          fontSize: {
            magnitude: 14,
            unit: "PT"
          },
          foregroundColor: {
            opaqueColor: {
              rgbColor: {
                blue: 1.0,
                green: 0.0,
                red: 0.0
              }
            }
          }
        },
        fields: "foregroundColor,fontFamily,fontSize"
      }
    },
    {
      updateTextStyle: {
        objectId: shapeId,
        textRange: {
          type: "FIXED_RANGE",
          startIndex: 10,
          endIndex: 15
        },
        style: {
          link: {
            url: "www.example.com"
          }
        },
        fields: "link"
      }
    }
  ];

  window.gapi.client.slides.presentations
    .batchUpdate({
      presentationId: presentationId,
      requests: requests
    })
    .then(batchUpdateResponse => {
      console.log(`Updated the text style for shape with ID: ${shapeId}`);
      callback(batchUpdateResponse.result);
    });
}

export function createBulletedText(presentationId, shapeId, callback) {
  var requests = [
    {
      createParagraphBullets: {
        objectId: shapeId,
        textRange: {
          type: "ALL"
        },
        bulletPreset: "BULLET_ARROW_DIAMOND_DISC"
      }
    }
  ];

  window.gapi.client.slides.presentations
    .batchUpdate({
      presentationId: presentationId,
      requests: requests
    })
    .then(batchUpdateResponse => {
      console.log(`Added bullets to text in shape with ID: ${shapeId}`);
      callback(batchUpdateResponse.result);
    });
}

export function createSheetsChart(
  presentationId,
  pageId,
  shapeId,
  sheetChartId,
  callback
) {
  var emu4M = {
    magnitude: 4000000,
    unit: "EMU"
  };
  var presentationChartId = "MyEmbeddedChart";
  var requests = [
    {
      createSheetsChart: {
        objectId: presentationChartId,
        spreadsheetId: shapeId,
        chartId: sheetChartId,
        linkingMode: "LINKED",
        elementProperties: {
          pageObjectId: pageId,
          size: {
            height: emu4M,
            width: emu4M
          },
          transform: {
            scaleX: 1,
            scaleY: 1,
            translateX: 100000,
            translateY: 100000,
            unit: "EMU"
          }
        }
      }
    }
  ];

  window.gapi.client.slides.presentations
    .batchUpdate({
      presentationId: presentationId,
      requests: requests
    })
    .then(batchUpdateResponse => {
      console.log(
        `Added a linked Sheets chart with ID: ${presentationChartId}`
      );
      callback(batchUpdateResponse.result);
    });
}

export function refreshSheetsChart(
  presentationId,
  presentationChartId,
  callback
) {
  var requests = [
    {
      refreshSheetsChart: {
        objectId: presentationChartId
      }
    }
  ];

  window.gapi.client.slides.presentations
    .batchUpdate({
      presentationId: presentationId,
      requests: requests
    })
    .then(batchUpdateResponse => {
      console.log(
        `Refreshed a linked Sheets chart with ID: ${presentationChartId}`
      );
      callback(batchUpdateResponse.result);
    });
}

export function addSlides(presentationId, num, layout, callback) {
  var requests = [];
  var slideIds = [];
  for (var i = 0; i < num; ++i) {
    slideIds.push(`slide_${i}`);
    requests.push({
      createSlide: {
        objectId: slideIds[i],
        slideLayoutReference: {
          predefinedLayout: layout
        }
      }
    });
  }
  var response = window.gapi.client.slides.presentations
    .batchUpdate({
      presentationId: presentationId,
      requests: requests
    })
    .then(response => {
      callback(slideIds);
    });
}

export function insertTextFromId(content, presentationId, objectID) {
  var requests = [
    {
      insertText: {
        objectId: objectID,
        insertionIndex: 0,
        text: content
      }
    }
  ];

  return requests;
  // window.gapi.client.slides.presentations.batchUpdate({
  //     presentationId: presentationId,
  //     requests: requests
  //   }).then((createTextboxWithTextResponse) => {
  //     console.log(`Created textbox with ID: ${createTextboxWithTextResponse}`);
  //   });
}

export function batchExecute(client, id, reqs, cbfunc) {
  client
    .batchUpdate({
      presentationId: id,
      requests: reqs
    })
    .then(response => {
      console.log(`batch execute : ${response}`);
      if (typeof cbfunc !== "undefined") cbfunc(response);
    });
}
