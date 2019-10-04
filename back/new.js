// Andrew Ha, 22246801



const officegen = require('officegen')
const fs = require('fs')

// Create an empty Word object:
let docx = officegen({
  type:'docx',
  pageMargins:{
    top:1000, right :1440, bottom:1400, left:1440,
  }
})




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

// uwa logo at top of page 193x64
let pObj = docx.createP()
pObj.addImage('uwalogo.jpg');



var interviewGuide = [
  [
  {
    val: "Interview Guide",
    opts: {
      b:true, // b = BOLD TEXT
      color: "FFFFFF",
      align: "center",
      vAlign: "center",
      cellColWidth: 9750,
      sz: '30',
      shd: {
        fill: "27348B",

        "themeFillTint": "80"
      }
    }
  }

],

  [{val: 'Appointment', opts: {b:true} }],
  [{val: 'Date', opts: {b:true} }],
  [{val: 'Interviewer', opts: {b:true} }],
  [{val: 'Candidate', opts: {b:true} }],


]

var interviewStyle = {
borderColor: '27348B',
tableColWidth: 0,
tableSize: 24,
tableColor: "27348B",
tableAlign: "left",
tableFontFamily: "calibri",
spacingBefor: 100, // default is 100
spacingAfter: 120, // default is 100
spacingLine: 240, // default is 240
spacingLineRule: 'atLeast', // default is atLeast
indent: -300, // table indent, default is 0
fixedLayout: true, // default is false
borders: true, // default is false. if true, default border size is 4
borderSize: 2, // To use this option, the 'borders' must set as true, default is 4
borderColor: "FFFFFF", // color for border XD
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
      color: 'FFFFFF',
      align: "center",
      vAlign: "center",
      shd: {
        fill: "27348B",

        "themeFillTint": "80"
      },

    }
  },
  {
    val: "Rating",
    opts: {
      b:true, // b = BOLD TEXT
      color: "FFFFFF",
      align: "center",
      vAlign: "center",
      cellColWidth: 1250,
      sz: '24',
      shd: {
        fill: "27348B",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "Met",
    opts: {
      align: "center",
      vAlign: "center",
      color: 'FFFFFF',
      cellColWidth: 2000,
      b:true,
      sz: '24',
      shd: {
        fill: "27348B",

        "themeFillTint": "80"
      }
    }
  }],

  ["1. Strategic Planning", '' , { val:meT, opts: {align: "center"} }],
   ["2. Leading Change", '' , { val:meT, opts: {align: "center"} }],
   ["3. Driving Execution", '' , { val:meT, opts: {align: "center"} }],
   ["4. Adaptability", '' , { val:meT, opts: {align: "center"} }],
   ["5. Initiation Action", "" , { val:meT, opts: {align: "center"} }],
   [testa, "" , { val:meT, opts: {align: "center"} }],


]

var tableStyle = {
  tableColWidth: 0,
  tableSize: 10,
  tableColor: "27348B",
  tableAlign: "left",
  tableFontFamily: "calibri",
  tableFontSize:20,
  spacingBefor: 50, // default is 100
  spacingAfter: 50, // default is 100
  spacingLine: 1, // default is 240
  spacingLineRule: 'atLeast', // default is atLeast
  indent: -300, // table indent, default is 0
  fixedLayout: true, // default is false
  borders: true, // default is false. if true, default border size is 4
  borderSize: 2, // To use this option, the 'borders' must set as true, default is 4

}
docx.createTable (table, tableStyle);


pObj = docx.createP();
pObj.addText('Final recommendation:            Appointable          Not Appointable',
 {bold: true, font_size:16, color: '27348B'})
 pObj.options.indentLeft = -250;


 var interviewNotes = [
   [
   {
     val: "Interview Notes",
     opts: {
       b:true, // b = BOLD TEXT
       color: "FFFFFF",
       align: "center",
       vAlign: "center",
       cellColWidth: 9750,
       sz: '24',
       shd: {
         fill: "27348B",

         "themeFillTint": "80"
       }
     }
   }

 ],
 ['\r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n']

]

var noteStyle = {
  tableColWidth: 0,
  tableSize: 16,
  tableColor: "27348B",
  tableAlign: "left",
  tableFontFamily: "calibri",
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


pObj = docx.createP(); // space between tables

var scale = [
  [{
    val: " ",
    opts: {
      cellColWidth: 400,
      b:true,
      sz: '24',
      align: "center",
      vAlign: "center",
      shd: {
        fill: "27348B",

        "themeFillTint": "80"
      },

    }
  },
  {
    val: "Rating",
    opts: {
      b:true, // b = BOLD TEXT
      color: "FFFFFF",
      align: "center",
      vAlign: "center",
      cellColWidth: 1500,
      sz: '24',
      shd: {
        fill: "27348B",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "Description",
    opts: {
      align: "center",
      vAlign: "center",
      color: 'FFFFFF',
      cellColWidth: 7850,
      b:true,
      sz: '24',
      shd: {
        fill: "27348B",

        "themeFillTint": "80"
      }
    }
  }],

  [{ val:'5', opts: {sz:20}}, { val:'Outstanding', opts: {sz:20}} , { val: 'Significantly exceeds criteria for successful job performance and organisational fit', opts: {sz:20}} ],
   [{ val:'4', opts: {sz:20}}, { val:'Excellent', opts: {sz:20}} , { val: "Exceeds criteria for successful job performance and organisational fit", opts: {sz:20}}],
   [{ val:'3', opts: {sz:20}}, { val:'Proficient' , opts: {sz:20}} , { val:  "Meets criteria for successful job performance and organisational fit",opts: {sz:20}}],
   [{ val:'2', opts: {sz:20}}, { val:'Basic', opts: {sz:20}} , { val: "Generally does not meet criteria for successful job performance and organisational fit",opts: {sz:20}}],
   [{ val:'1', opts: {sz:20}}, { val:"Limited", opts: {sz:20}} , { val: "Significantly below criteria required for successful job performance and organisational fit",opts: {sz:20}}],



]

var scaleStyle = {
  tableColWidth: 0,
  tableSize: 10,
  tableColor: "000000",
  tableAlign: "center",
  tableFontFamily: "calibri",
  spacingBefor: 50, // default is 100
  spacingAfter: 50, // default is 100
  spacingLine: 25, // default is 240
  spacingLineRule: 'atLeast', // default is atLeast
  indent: -300, // table indent, default is 0
  fixedLayout: true, // default is false
  borders: true, // default is false. if true, default border size is 4
  borderSize: 2, // To use this option, the 'borders' must set as true, default is 4
}
docx.createTable (scale, scaleStyle);


// END OF PAGE 1
// START OF PAGE 2


pObj = docx.createP()
pObj.addImage('uwalogo.jpg');


// opening table
var openingTable = [
  [{
    val: "Opening",
    opts: {
      cellColWidth: 9750,
      color: "FFFFFF",
      b:true,
      sz: '30',
      align: "center",
      vAlign: "center",
      shd: {
        fill: "27348B",

        "themeFillTint": "80"
      },

    }
  }

  ],

]

var scaleStyle2 = {
  tableColWidth: 0,
  tableSize: 10,
  tableColor: "FFFFFF",
  tableAlign: "center",
  tableFontFamily: "calibri",
  spacingBefor: 50, // default is 100
  spacingAfter: 50, // default is 100
  spacingLine: 25, // default is 240
  spacingLineRule: 'atLeast', // default is atLeast
  indent: -300, // table indent, default is 0
  fixedLayout: true, // default is false
  borders: true, // default is false. if true, default border size is 4
  borderSize: 2, // To use this option, the 'borders' must set as true, default is 4
}

docx.createTable (openingTable, scaleStyle2);


var whydidyou = [
  [{
    val: "Why did you apply?",
    opts: {
      cellColWidth: 9750,
      b:false,
      color: "27348B",
      sz: '24',
      align: "left",
      vAlign: "center",
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      },

    }
  }

  ],

]

var scaleStyle3 = {
  tableColWidth: 0,
  tableSize: 10,
  tableColor: "27348B",
  tableAlign: "center",
  tableFontFamily: "calibri",
  spacingBefor: 50, // default is 100
  spacingAfter: 50, // default is 100
  spacingLine: 25, // default is 240
  spacingLineRule: 'atLeast', // default is atLeast
  indent: -300, // table indent, default is 0
  fixedLayout: true, // default is false
  borders: false, // default is false. if true, default border size is 4
  borderSize: 2, // To use this option, the 'borders' must set as true, default is 4k
}


docx.createTable (whydidyou, scaleStyle3);



var blankBox = [
  [{
    val: "\r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n",
    opts: {
      cellColWidth: 9750,
      b:false,
      color: "27348B",
      sz: '20',
      align: "left",
      vAlign: "center",
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      },

    }
  }

  ],

]

var scaleStyle4 = {
  tableColWidth: 0,
  tableSize: 10,
  tableColor: "FFFFFF",
  tableAlign: "center",
  tableFontFamily: "calibri",
  spacingBefor: 50, // default is 100
  spacingAfter: 50, // default is 100
  spacingLine: 25, // default is 240
  spacingLineRule: 'atLeast', // default is atLeast
  indent: -300, // table indent, default is 0
  fixedLayout: true, // default is false
  borders: true, // default is false. if true, default border size is 4
  borderSize: 2, // To use this option, the 'borders' must set as true, default is 4
}


docx.createTable (blankBox, scaleStyle4);


var whatdoyoulike = [
  [{
    val: "What do you like most about your current role, what do you like the least and why are you looking to leave?",
    opts: {
      cellColWidth: 9750,
      b:false,
      color: "27348B",
      sz: '21',
      align: "left",
      vAlign: "center",
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      },

    }
  }

  ],
  [{val: "Most: _____________________________________________________________", opts: {align:"right"}}],
  [{val: "Least: _____________________________________________________________", opts: {align:"right"}}],
  [{val: "Reason for Leaving: _____________________________________________________________", opts: {align:"right" }}],
  [{val:"What are your career objectives?", opts: {align:"left"}}]

]

docx.createTable(whatdoyoulike, scaleStyle3);

docx.createTable (blankBox, scaleStyle4);
docx.putPageBreak();


// END OF PAGE 2
// START OF PAGE 3

pObj = docx.createP()
pObj.addImage('uwalogo.jpg');


var behavAssessment = [
  [
    {
    val: "",
      opts: {
      cellColWidth: 4875,
      b:false,
      color: "27348B",
      sz: '24',
      align: "left",
      vAlign: "center",
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      },

    }
  },
  {
  val: "",
    opts: {
    cellColWidth: 4875,
    b:false,
    color: "27348B",
    sz: '24',
    align: "left",
    vAlign: "center",
    shd: {
      fill: "FFFFFF",

      "themeFillTint": "80"
    },

  }
},



  ],
  ["Leading Change: Driving organizational and cultural changes needed to achieve strategic objectives catalysing new approaches to improve results by transforming organizational culture, sustems or products/services; helping others overcome resistance to change.",
  ["Key Actions:",
  "Gathers information-Identifies and fills gaps in information required to understand business issues and opportunities.",
  "Organises information-Organizes information and data to identifty major trends and problems; comapres and combines information to understand underlying issues and predict future trends",
  "Evaluates/Selects Strategies-Generates and cosiders options for action to achieve a long-range goal; develops decvision criteria cosidering factors such as cost, benefits risks, timing and buy-in; selects the strategy most likely to succeed.",
  "Establishes high-level plan-Identifies the key tasks and resources needed to achieve strategic objectives."
  ]
  ],


]

var scaleStyle5 = {
  tableColWidth: 0,
  tableSize: 10,
  tableColor: "27348B",
  tableAlign: "left",
  tableFontFamily: "calibri",
  spacingBefor: 50, // default is 100
  spacingAfter: 50, // default is 100
  spacingLine: 25, // default is 240
  spacingLineRule: 'atLeast', // default is atLeast
  indent: -300, // table indent, default is 0
  fixedLayout: true, // default is false
  borders: false, // default is false. if true, default border size is 4
  borderSize: 2, // To use this option, the 'borders' must set as true, default is 4k
}


docx.createTable (behavAssessment, scaleStyle5);


pObj = docx.createP();
pObj.addText('What strategies have you employed to ensure a major new directive from senior management was carried out?', { color: '000088', font_face: 'calibri', font_size: 12, indent:-300 })























// Let's generate the Word document into a file:

let out = fs.createWriteStream('InterviewGuide.docx')

out.on('error', function(err) {
  console.log(err)
})

// Async call to generate the output file:
docx.generate(out)
