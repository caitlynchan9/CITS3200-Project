const officegen = require('officegen')
const fs = require('fs')

// Create an empty Word object:
let docx = officegen('docx')

// Officegen calling this function after finishing to generate the docx document:
docx.on('finalize', function(written) {
  console.log(
    'Finish to create a Microsoft Word document.'
  )
})

// Officegen calling this function to report errors:
docx.on('error', function(err) {
  console.log(err)
})

var testa = "6. Technology Savvy";
var meT = "Yes / No"

// Create a new paragraph:
let pObj = docx.createP()
pObj.addImage('uwalogo.jpg');



var interviewGuide = [
  [
  {
    val: "Interview Guide",
    opts: {
      b:true, // b = BOLD TEXT
      color: "A00000",
      align: "center",
      vAlign: "center",
      cellColWidth: 9750,
      sz: '36',
      shd: {
        fill: "92CDDC",
        themeFill: "text1",
        "themeFillTint": "80"
      }
    }
  }

],

  ['Appointment',],
  ['Date'],
  ['Interviewer'],
  ['Candidate'],


]

var interviewStyle = {
tableColWidth: 0,
tableSize: 24,
tableColor: "ada",
tableAlign: "left",
tableFontFamily: "Times New Roman",
spacingBefor: 100, // default is 100
spacingAfter: 120, // default is 100
spacingLine: 240, // default is 240
spacingLineRule: 'atLeast', // default is atLeast
indent: -300, // table indent, default is 0
fixedLayout: true, // default is false
borders: true, // default is false. if true, default border size is 4
borderSize: 2, // To use this option, the 'borders' must set as true, default is 4
}

docx.createTable (interviewGuide, interviewStyle);

pObj = docx.createP(); // creating space between table



// create a table for the interview guide

var table = [
  [{
    val: "Selection Criteria",
    opts: {
      cellColWidth: 6500,
      b:true,
      sz: '24',
      align: "center",
      vAlign: "center",
      shd: {
        fill: "7F7F7F",
        themeFill: "text1",
        "themeFillTint": "80"
      },
      fontFamily: "Avenir Book"
    }
  },
  {
    val: "Rating",
    opts: {
      b:true, // b = BOLD TEXT
      color: "A00000",
      align: "center",
      vAlign: "center",
      cellColWidth: 1250,
      sz: '24',
      shd: {
        fill: "92CDDC",
        themeFill: "text1",
        "themeFillTint": "80"
      }
    }
  },
  {
    val: "Met",
    opts: {
      align: "center",
      vAlign: "center",
      cellColWidth: 2000,
      b:true,
      sz: '24',
      shd: {
        fill: "92CDDC",
        themeFill: "text1",
        "themeFillTint": "80"
      }
    }
  }],

  ["1. Strategic Planning", '' , meT ],
   ["2. Leading Change", '' , meT],
   ["3. Driving Execution", '' , meT],
   ["4. Adaptability", '' , meT],
   ["5. Initiation Action", "" , meT],
   [testa, "" , meT],


]

var tableStyle = {
  tableColWidth: 0,
  tableSize: 10,
  tableColor: "ada",
  tableAlign: "left",
  tableFontFamily: "Times New Roman",
  tableFontSize:20,
  spacingBefor: 0, // default is 100
  spacingAfter: 0, // default is 100
  spacingLine: 25, // default is 240
  spacingLineRule: 'atLeast', // default is atLeast
  indent: -300, // table indent, default is 0
  fixedLayout: true, // default is false
  borders: true, // default is false. if true, default border size is 4
  borderSize: 2, // To use this option, the 'borders' must set as true, default is 4
}
docx.createTable (table, tableStyle);


pObj = docx.createP();
pObj.addText('Final recommendation:            Appointable          Not Appointable',
 {bold: true, spacingBefor: 200, font_size:16})
 pObj.options.indentLeft = -250;


 var interviewNotes = [
   [
   {
     val: "Interview Notes",
     opts: {
       b:true, // b = BOLD TEXT
       color: "A00000",
       align: "center",
       vAlign: "center",
       cellColWidth: 9750,
       sz: '24',
       shd: {
         fill: "92CDDC",
         themeFill: "text1",
         "themeFillTint": "80"
       }
     }
   }

 ],
 ['\r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n']

]

var noteStyle = {
  tableColWidth: 0,
  tableSize: 16,
  tableColor: "ada",
  tableAlign: "left",
  tableFontFamily: "Times New Roman",
  spacingBefor: 100, // default is 100
  spacingAfter: 120, // default is 100
  spacingLine: 240, // default is 240
  spacingLineRule: 'atLeast', // default is atLeast
  indent: -300, // table indent, default is 0
  fixedLayout: true, // default is false
  borders: true, // default is false. if true, default border size is 4
  borderSize: 2, // To use this option, the 'borders' must set as true, default is 4
}

docx.createTable(interviewNotes,noteStyle);


var scale = [
  [{
    val: " ",
    opts: {
      cellColWidth: 300,
      b:true,
      sz: '24',
      align: "center",
      vAlign: "center",
      shd: {
        fill: "7F7F7F",
        themeFill: "text1",
        "themeFillTint": "80"
      },
      fontFamily: "Avenir Book"
    }
  },
  {
    val: "Rating",
    opts: {
      b:true, // b = BOLD TEXT
      color: "A00000",
      align: "center",
      vAlign: "center",
      cellColWidth: 1500,
      sz: '24',
      shd: {
        fill: "92CDDC",
        themeFill: "text1",
        "themeFillTint": "80"
      }
    }
  },
  {
    val: "Met",
    opts: {
      align: "center",
      vAlign: "center",
      cellColWidth: 7950,
      b:true,
      sz: '24',
      shd: {
        fill: "92CDDC",
        themeFill: "text1",
        "themeFillTint": "80"
      }
    }
  }],

  ["5", 'Outstanding' , "Significantly exceeds criteria for successful job performance and organisational fit" ],
   ["4", 'Excellent' , "Exceeds criteria for successful job performance and organisational fit"],
   ["3", 'Proficient' , "Meets criteria for successful job performance and organisational fit"],
   ["2", 'Basic' , "Generally does not meet criteria for successful job performance and organisational fit"],
   ["1", "Limited" , "Significantly below criteria required for successful job performance and organisational fit"],



]

var scaleStyle = {
  tableColWidth: 0,
  tableSize: 10,
  tableColor: "ada",
  tableAlign: "center",
  tableFontFamily: "Times New Roman",
  font_size:12,
  spacingBefor: 0, // default is 100
  spacingAfter: 0, // default is 100
  spacingLine: 25, // default is 240
  spacingLineRule: 'atLeast', // default is atLeast
  indent: -300, // table indent, default is 0
  fixedLayout: true, // default is false
  borders: true, // default is false. if true, default border size is 4
  borderSize: 2, // To use this option, the 'borders' must set as true, default is 4
}
docx.createTable (scale, scaleStyle);

















// Let's generate the Word document into a file:

let out = fs.createWriteStream('example.docx')

out.on('error', function(err) {
  console.log(err)
})

// Async call to generate the output file:
docx.generate(out)
