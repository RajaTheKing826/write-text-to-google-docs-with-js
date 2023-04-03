const express = require("express")
const docs = require("@googleapis/docs")

const documentId = "1AmeNmAvgTXrrvCXRv9e9_6zjcwJEBfYoGVehUwXznK0"
const rangeName = "webFlowScriptTags"
async function getGoogleData() {
  const auth = new docs.auth.GoogleAuth({
    keyFile: "credentials.json",
    scopes: ["https://www.googleapis.com/auth/documents"],
  })
  const authClient = await auth.getClient()

  const client = await docs.docs({
    version: "v1",
    auth: authClient,
  })

  let requests = []

  const newText = "Some Re Updated text"

  const newTextLength = newText.length

  const document = await client.documents.get({
    documentId: documentId,
  })

  const namedRanges = await document.namedRanges

  if (namedRanges) {
    const filteredNamedRangesList = await document.namedRanges[rangeName]
    if (!filteredNamedRangesList) {
      console.log("empty named range")
    } else {
      let allRanges = []
      let insertAt = {}
      filteredNamedRangesList.namedRanges.forEach((namedRange) => {
        let ranges = namedRange.ranges
        allRanges.push(ranges)
        insertAt[ranges[0].startIndex] = true
      })

      function compare(a, b) {
        if (a.startIndex > b.startIndex) {
          return -1
        }
        if (a.startIndex < b.startIndex) {
          return 1
        }
        return 0
      }

      allRanges.sort(compare)
      allRanges.forEach((rangeValue) => {
        requests.push({
          deleteContentRange: {
            range: rangeValue,
          },
        })

        const segmentId = rangeValue.get("segmentId")
        const start = rangeValue.get("startIndex")

        if (insertAt[start]) {
          requests.push({
            insertText: {
              location: {
                segmentId: segmentId,
                index: start,
              },
              text: newText,
            },
          })
          requests.push({
            createNamedRange: {
              name: rangeName,
              range: {
                segmentId: segmentId,
                startIndex: start,
                endIndex: start + new_text_len,
              },
            },
          })
        }
      })
    }
  } else {
    requests.push({
      insertText: {
        location: {
          index: 1,
          segmentId: "",
        },
        text: newText,
      },
    })
    requests.push({
      createNamedRange: {
        name: rangeName,
        range: {
          startIndex: 1,
          endIndex: newTextLength,
          segmentId: "",
        },
      },
    })
  }
  const res = await client.documents.batchUpdate({
    documentId: documentId,
    requestBody: {
      requests: requests,
    },
  })
  console.log(res.data)
}

getGoogleData().catch((e) => {
  console.error(e)
  throw e
})
