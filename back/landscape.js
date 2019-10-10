const officegen = require('officegen');
const fs = require('fs');

let docx = officegen(
  {
    type:'docx',
    orientation:'landscape',
    pageMargins:{
      top:800, right :1440, bottom:400, left:1440,
    }
  });


// Officegen calling this function after finishing to generate the docx document:
docx.on('finalize', function(written)
  {
    console.log('Finish to create a Microsoft Word document.')
  })

// Officegen calling this function to report errors:
docx.on('error', function(err)
  {
    console.log(err)
  })




let pObj = docx.createP({align:'center'})
pObj.addImage('uwalogo.jpg');



var assessmentMatrix = [
  [
    {
      val:"Assessment Matrix",
      opts:
      {
        cellColWidth:15250,
        color:'FFFFFF',
        b:true,
        sz:'30',
        align:'center',
        vAlign:'center',
        shd:
        {
          fill:"27348B"
        }

      }
    }
  ],

  [{val:'Job Title / Job Number', opts:{b:true}}],
  [{val:'Selection Panel Members', opts:{b:true}}],
  [{val:'Interview Date(s)', opts:{b:true}}],




]


var matrixStyle = {
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
  borderColor: "27348B", // color for border XD doesnt work
}

docx.createTable(assessmentMatrix, matrixStyle);




pObj = docx.createP() // space between tables

var selectionCriteria = [
  [{
    val: "Selection Criteria",
    opts: {
      cellColWidth: 12000,
      color: "FFFFFF",
      b:true,
      sz: '30',
      align: "center",
      vAlign: "center",
      shd: {
        fill: "27348B",

      },

    }
  },
  {
    val:"Candidate Ranking",
    opts:{
      cellColWidth:3250,
      color:'FFFFFF',
      b:true,
      sz:'30',
      align: 'center',
      vAlign: 'center',
      shd:{
        fill:"27348B",

      }
    }
  }

  ],


]

var selectStyle = {
  tableColWidth: 0,
  tableSize: 10,
  tableColor: "27348B",
  tableAlign: "left",
  tableFontFamily: "calibri",
  spacingBefor: 100, // default is 100
  spacingAfter: 120, // default is 100
  spacingLine: 25, // default is 240
  spacingLineRule: 'atLeast', // default is atLeast
  indent: -300, // table indent, default is 0
  fixedLayout: true, // default is false
  borders: true, // default is false. if true, default border size is 4
  borderSize: 2, // To use this option, the 'borders' must set as true, default is 4
}

docx.createTable (selectionCriteria, selectStyle);



var scale = [
  [{
    val: "Selection Criteria",
    opts: {
      cellColWidth: 3000,
      color: '27348B',
      b:true,
      sz: '24',
      align: "center",
      vAlign: "center",
      shd: {
        fill: "CCCCCC",

        "themeFillTint": "80"
      },

    }
  },
  {
    val: "",
    opts: {
      b:true, // b = BOLD TEXT
      color: "27348B",
      align: "center",
      vAlign: "center",
      cellColWidth: 1800,
      sz: '24',
      shd: {
        fill: "CCCCCC",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "",
    opts: {
      align: "center",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1800,
      b:true,
      sz: '24',
      shd: {
        fill: "CCCCCC",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "",
    opts: {
      align: "center",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1800,
      b:true,
      sz: '24',
      shd: {
        fill: "CCCCCC",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "",
    opts: {
      align: "center",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1800,
      b:true,
      sz: '24',
      shd: {
        fill: "CCCCCC",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "",
    opts: {
      align: "center",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1800,
      b:true,
      sz: '24',
      shd: {
        fill: "CCCCCC",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "Ranking & Candidate Names",
    opts: {
      align: "center",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 3250,
      b:true,
      sz: '24',
      shd: {
        fill: "CCCCCC",

        "themeFillTint": "80"
      }
    }
  }

],

  [ { val:'Strategic Planning and Organizing', opts: {sz:20}}, "", "" , "", "", "", {val:'1st', opts:{b:true}}],
  [ { val:'Leading Change', opts: {sz:20}}, "", "" , "", "", "", {val:'2nd', opts:{b:true}} ],
  [ { val:'Driving Execution', opts: {sz:20}}, "", "" , "", "", "", {val:'3rd', opts:{b:true}}],
  [ { val:'Adaptability', opts: {sz:20}}, "", "" , "", "", "", {val:'4th', opts:{b:true}}],
  [ { val:'Initiating Action', opts: {sz:20}}, "", "" , "", "", "", {val:'5th', opts:{b:true}} ],
  [ { val:'Technology Savvy', opts: {sz:20}}, "", "" , "", "", "", "" ],



]

var scaleStyle = {
  tableColWidth: 0,
  tableSize: 10,
  tableColor: "27348B",
  tableAlign: "left",
  tableFontFamily: "calibri",
  spacingBefor: 100, // default is 100
  spacingAfter: 100, // default is 100
  spacingLine: 25, // default is 240
  spacingLineRule: 'atLeast', // default is atLeast
  indent: -300, // table indent, default is 0
  fixedLayout: true, // default is false
  borders: true, // default is false. if true, default border size is 4
  borderSize: 2, // To use this option, the 'borders' must set as true, default is 4
}
docx.createTable (scale, scaleStyle);



var appointableNA = [
  [{
    val: ["Appointable(A) or", "Not Appointable(NA)"],
    opts: {
      cellColWidth: 3000,
      color: '27348B',
      b:true,
      sz: '20',
      align: "left",

      shd: {
        fill: "CCCCCC",

        "themeFillTint": "80"
      },

    }
  },
  {
    val: "A                 NA",
    opts: {
      b:true, // b = BOLD TEXT
      color: "27348B",
      align: "left",
      vAlign: "center",
      cellColWidth: 1800,
      sz: '24',
      shd: {
        fill: "CCCCCC",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "A                 NA",
    opts: {
      align: "left",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1800,
      b:true,
      sz: '24',
      shd: {
        fill: "CCCCCC",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "A                 NA",
    opts: {
      align: "left",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1800,
      b:true,
      sz: '24',
      shd: {
        fill: "CCCCCC",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "A                 NA",
    opts: {
      align: "left",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1800,
      b:true,
      sz: '24',
      shd: {
        fill: "CCCCCC",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "A                 NA",
    opts: {
      align: "left",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1800,
      b:true,
      sz: '24',
      shd: {
        fill: "CCCCCC",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "",
    opts: {
      align: "left",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 3250,
      b:true,
      sz: '24',
      shd: {
        fill: "CCCCCC",

        "themeFillTint": "80"
      }
    }
  }

],

]

var appoStyle = {
  tableColWidth: 0,
  tableSize: 10,
  tableColor: "27348B",
  tableAlign: "left",
  tableFontFamily: "calibri",
  spacingBefor: 100, // default is 100
  spacingAfter: 100, // default is 100
  spacingLine: 25, // default is 240
  spacingLineRule: 'atLeast', // default is atLeast
  indent: -300, // table indent, default is 0
  fixedLayout: true, // default is false
  borders: true, // default is false. if true, default border size is 4
  borderSize: 2, // To use this option, the 'borders' must set as true, default is 4
}
docx.createTable (appointableNA, appoStyle);



var notifyC = [
  [{
    val: ["Who will notify the candidate?", "(Panel Chair/Talent Acquisition)"],
    opts: {
      cellColWidth: 3000,
      color: '27348B',
      b:true,
      sz: '20',
      align: "left",

      shd: {
        fill: "CCCCCC",

        "themeFillTint": "80"
      },

    }
  },
  {
    val: "PC                TA",
    opts: {
      b:true, // b = BOLD TEXT
      color: "27348B",
      align: "left",
      vAlign: "center",
      cellColWidth: 1800,
      sz: '24',
      shd: {
        fill: "CCCCCC",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "PC                TA",
    opts: {
      align: "left",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1800,
      b:true,
      sz: '24',
      shd: {
        fill: "CCCCCC",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "PC                TA",
    opts: {
      align: "left",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1800,
      b:true,
      sz: '24',
      shd: {
        fill: "CCCCCC",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "PC                TA",
    opts: {
      align: "left",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1800,
      b:true,
      sz: '24',
      shd: {
        fill: "CCCCCC",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "PC                TA",
    opts: {
      align: "left",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1800,
      b:true,
      sz: '24',
      shd: {
        fill: "CCCCCC",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "",
    opts: {
      align: "left",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 3250,
      b:true,
      sz: '24',
      shd: {
        fill: "CCCCCC",

        "themeFillTint": "80"
      }
    }
  }

],

]

var appoStyle = {
  tableColWidth: 0,
  tableSize: 10,
  tableColor: "27348B",
  tableAlign: "left",
  tableFontFamily: "calibri",
  spacingBefor: 100, // default is 100
  spacingAfter: 100, // default is 100
  spacingLine: 25, // default is 240
  spacingLineRule: 'atLeast', // default is atLeast
  indent: -300, // table indent, default is 0
  fixedLayout: true, // default is false
  borders: true, // default is false. if true, default border size is 4
  borderSize: 2, // To use this option, the 'borders' must set as true, default is 4
}
docx.createTable (notifyC, appoStyle);

docx.putPageBreak();

// END OF MATRIX PAGE




pObj = docx.createP({align:'center'})
pObj.addImage('uwalogo.jpg');



var panelRecommendation= [
  [
    {
      val:"Panel Recommendation",
      opts:
      {
        cellColWidth:15250,
        color:'FFFFFF',
        b:true,
        sz:'30',
        align:'center',
        vAlign:'center',
        shd:
        {
          fill:"27348B"
        }

      }
    }
  ],

]


var matrixStyle = {
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
  borderColor: "27348B", // color for border XD doesnt work
}

docx.createTable(panelRecommendation, matrixStyle);



var blankBox = [
  [{
    val: "Please provide a brief summary of the recommendation(s) with reference to interview performance, panel discussion etc  \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n ",
    opts: {
      cellColWidth: 15250,
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

var blankStyle = {
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


docx.createTable (blankBox, blankStyle);


var appointBox = [
  [{
    val: [
    "The panel unanimously agreed to appoint ___________________ to the positions. " +
    "This applicant exceeded the selection criteria for the position in their job application, " +
    "interview and reference checks. The suggested offer is as follows:" , "",
    "Level:                                 ______________________________",
    "Step:                                  ______________________________",
    "Direct Supervisor:           ______________________________",
    "Allowance(s):                   ______________________________",
    "FTE:                                    ______________________________",
    "Estimated Start Date:     ______________________________",
    "End Date (if applicable): ______________________________",
    "Visa Sponsorship:            ______________________________",
    "Relocation Support:        ______________________________",
    "Comments:                       ______________________________",


    ]
    ,
      opts: {
      cellColWidth: 15250,
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

var blankStyle = {
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


docx.createTable (appointBox, blankStyle);












// Let's generate the Word document into a file:

let out = fs.createWriteStream('Matrix.docx')

out.on('error', function(err) {
  console.log(err)
})

// Async call to generate the output file:
docx.generate(out);
