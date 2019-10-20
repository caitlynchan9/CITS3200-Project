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
  tableColor: "27348B",
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
  ["Strategic Planning: Obtaining information and identifying key issues and relationships relevant to achieving a long-range goal; committing to a course of action to accomplish a long-range goal after developing alternatives based on logical assumptions, facts, available resources, constraints and organizational values.",
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


pObj = docx.createP();
pObj.addText('What strategies have you employed to ensure a major new directive from senior management was carried out?', { color: '000088', font_face: 'calibri', font_size: 12, indent:-300 })



var openingTable = [
  [{
    val: "Response",
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


var blankBox = [
  [{
    val: "\r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n ",
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



var scale = [
  [{
    val: "0",
    opts: {
      cellColWidth: 1625,
      color: '27348B',
      b:true,
      sz: '24',
      align: "center",
      vAlign: "center",
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      },

    }
  },
  {
    val: "1",
    opts: {
      b:true, // b = BOLD TEXT
      color: "27348B",
      align: "center",
      vAlign: "center",
      cellColWidth: 1625,
      sz: '24',
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "2",
    opts: {
      align: "center",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1625,
      b:true,
      sz: '24',
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "3",
    opts: {
      align: "center",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1625,
      b:true,
      sz: '24',
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "4",
    opts: {
      align: "center",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1625,
      b:true,
      sz: '24',
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "5",
    opts: {
      align: "center",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1625,
      b:true,
      sz: '24',
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      }
    }
  }
],

  [ { val:'Does Not Meet', opts: {sz:20}}, { val:'Limited', opts: {sz:20}}, { val:'Basic', opts: {sz:20}},
  { val:'Proficient', opts: {sz:20}},
{ val:'Excellent', opts: {sz:20}},{ val:'Outstanding', opts: {sz:20}}, ],



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


docx.putPageBreak();


//END OF PAGE 3

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
  "Identifies Change Opportunities-Proactively recognises a need and takes accountability for implementing an improvement and/or change; looks for opportunities to mobilize others to implement new solutions.",
  "Catalyses Change-Creates momentum by taking immediate ation and encouraging others to take action to improve organisational culture, processes or products/services; offers resources and direction to support implementation; breaks down cultural and operational barriers to change; recognises and rewards those who contribute to change efforts.",
  ,
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


pObj = docx.createP();
pObj.addText('What strategies have you employed to ensure a major new directive from senior management was carried out?', { color: '000088', font_face: 'calibri', font_size: 12, indent:-300 })



var openingTable = [
  [{
    val: "Response",
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


var blankBox = [
  [{
    val: "\r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n ",
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



var scale = [
  [{
    val: "0",
    opts: {
      cellColWidth: 1625,
      color: '27348B',
      b:true,
      sz: '24',
      align: "center",
      vAlign: "center",
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      },

    }
  },
  {
    val: "1",
    opts: {
      b:true, // b = BOLD TEXT
      color: "27348B",
      align: "center",
      vAlign: "center",
      cellColWidth: 1625,
      sz: '24',
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "2",
    opts: {
      align: "center",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1625,
      b:true,
      sz: '24',
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "3",
    opts: {
      align: "center",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1625,
      b:true,
      sz: '24',
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "4",
    opts: {
      align: "center",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1625,
      b:true,
      sz: '24',
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "5",
    opts: {
      align: "center",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1625,
      b:true,
      sz: '24',
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      }
    }
  }
],

  [ { val:'Does Not Meet', opts: {sz:20}}, { val:'Limited', opts: {sz:20}}, { val:'Basic', opts: {sz:20}},
  { val:'Proficient', opts: {sz:20}},
{ val:'Excellent', opts: {sz:20}},{ val:'Outstanding', opts: {sz:20}}, ],



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


docx.putPageBreak();

// END OF PAGE 4

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
  ["Driving Execution: Translating strategic priorities into operational reality; aligning communication, accountabilities, resource capabilities, internal processes, and ongoing measurement systems to ensure that strategic priorities yield measurable and sustainable results.",
  ["Key Actions:",
  "Translates initiatives into actions—Determines action steps and milestones required to implement a specific business initiative; adjusts activities or timelines as circumstances warrant.",
  "Communicates to engage others—Establishes two-way communication channels to convey business strategies and plans; engages people by helping them understand the reasons behind organizational initiatives and the value of assigned responsibilities for the individual, team, and organization.",
  "Measures progress—Establishes criteria and systems (including lead and lag measures) to track ongoing progress toward goals; follows up on assigned responsibilities."
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


pObj = docx.createP();
pObj.addText('What strategies have you employed to ensure a major new directive from senior management was carried out?', { color: '000088', font_face: 'calibri', font_size: 12, indent:-300 })



var openingTable = [
  [{
    val: "Response",
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


var blankBox = [
  [{
    val: "\r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n ",
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



var scale = [
  [{
    val: "0",
    opts: {
      cellColWidth: 1625,
      color: '27348B',
      b:true,
      sz: '24',
      align: "center",
      vAlign: "center",
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      },

    }
  },
  {
    val: "1",
    opts: {
      b:true, // b = BOLD TEXT
      color: "27348B",
      align: "center",
      vAlign: "center",
      cellColWidth: 1625,
      sz: '24',
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "2",
    opts: {
      align: "center",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1625,
      b:true,
      sz: '24',
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "3",
    opts: {
      align: "center",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1625,
      b:true,
      sz: '24',
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "4",
    opts: {
      align: "center",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1625,
      b:true,
      sz: '24',
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "5",
    opts: {
      align: "center",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1625,
      b:true,
      sz: '24',
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      }
    }
  }
],

  [ { val:'Does Not Meet', opts: {sz:20}}, { val:'Limited', opts: {sz:20}}, { val:'Basic', opts: {sz:20}},
  { val:'Proficient', opts: {sz:20}},
{ val:'Excellent', opts: {sz:20}},{ val:'Outstanding', opts: {sz:20}}, ],



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


docx.putPageBreak();

// END OF PAGE 5

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
  ["Adaptability: Maintaining effectiveness when experiencing major changes in work responsibilities or environment (e.g., people, processes, structure, or culture); adjusting effectively to change by exploring the benefits, trying new approaches, and collaborating with others to make the change successful.",
  ["Key Actions:",
  "Tries to understand changes—Actively seeks information (from coworkers, leaders, customers, competition, technologies, and regulations) to understand the rationale and implications for changes.",
  "Approaches change with a positive mind-set—Treats new situations as opportunities for learning or growth; actively seeks to identify and communicate the benefits of changes; collaborates with others to implement changes.",
  "Adjusts behavior—Quickly modifies daily behavior and tries new approaches to deal effectively with changes; does not persist with ineffective methods; leverages available resources to ease transition."
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


pObj = docx.createP();
pObj.addText('What strategies have you employed to ensure a major new directive from senior management was carried out?', { color: '000088', font_face: 'calibri', font_size: 12, indent:-300 })



var openingTable = [
  [{
    val: "Response",
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


var blankBox = [
  [{
    val: "\r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n ",
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



var scale = [
  [{
    val: "0",
    opts: {
      cellColWidth: 1625,
      color: '27348B',
      b:true,
      sz: '24',
      align: "center",
      vAlign: "center",
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      },

    }
  },
  {
    val: "1",
    opts: {
      b:true, // b = BOLD TEXT
      color: "27348B",
      align: "center",
      vAlign: "center",
      cellColWidth: 1625,
      sz: '24',
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "2",
    opts: {
      align: "center",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1625,
      b:true,
      sz: '24',
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "3",
    opts: {
      align: "center",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1625,
      b:true,
      sz: '24',
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "4",
    opts: {
      align: "center",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1625,
      b:true,
      sz: '24',
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "5",
    opts: {
      align: "center",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1625,
      b:true,
      sz: '24',
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      }
    }
  }
],

  [ { val:'Does Not Meet', opts: {sz:20}}, { val:'Limited', opts: {sz:20}}, { val:'Basic', opts: {sz:20}},
  { val:'Proficient', opts: {sz:20}},
{ val:'Excellent', opts: {sz:20}},{ val:'Outstanding', opts: {sz:20}}, ],



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


docx.putPageBreak();

// END OF PAGE 6

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
  ["Initiating Action: Taking prompt action to accomplish work goals; taking action to achieve results beyond what is required; being proactive.",
  ["Key Actions:",
  "Responds quickly—Takes immediate action when confronted with a problem or when made aware of a situation.",
  "Takes independent action—Implements new ideas or potential solutions without prompting; does not wait for others to take action or to request action.",
  "Goes above and beyond—Takes action that goes beyond job requirements in order to achieve results."
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


pObj = docx.createP();
pObj.addText('What strategies have you employed to ensure a major new directive from senior management was carried out?', { color: '000088', font_face: 'calibri', font_size: 12, indent:-300 })



var openingTable = [
  [{
    val: "Response",
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


var blankBox = [
  [{
    val: "\r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n ",
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



var scale = [
  [{
    val: "0",
    opts: {
      cellColWidth: 1625,
      color: '27348B',
      b:true,
      sz: '24',
      align: "center",
      vAlign: "center",
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      },

    }
  },
  {
    val: "1",
    opts: {
      b:true, // b = BOLD TEXT
      color: "27348B",
      align: "center",
      vAlign: "center",
      cellColWidth: 1625,
      sz: '24',
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "2",
    opts: {
      align: "center",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1625,
      b:true,
      sz: '24',
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "3",
    opts: {
      align: "center",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1625,
      b:true,
      sz: '24',
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "4",
    opts: {
      align: "center",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1625,
      b:true,
      sz: '24',
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "5",
    opts: {
      align: "center",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1625,
      b:true,
      sz: '24',
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      }
    }
  }
],

  [ { val:'Does Not Meet', opts: {sz:20}}, { val:'Limited', opts: {sz:20}}, { val:'Basic', opts: {sz:20}},
  { val:'Proficient', opts: {sz:20}},
{ val:'Excellent', opts: {sz:20}},{ val:'Outstanding', opts: {sz:20}}, ],



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


docx.putPageBreak();

//END OF PAGE 7

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
  ["Technology Savvy: Leveraging one’s practical knowledge and understanding of recent technology tools, solutions, and trends to improve work results, solve work problems, and take advantage of new business opportunities.",
  ["Key Actions:",
  "Actively develops expertise—Pursues opportunities to develop knowledge and experiment with latest technology solutions that can help accomplish work goals; when necessary, overcomes own resistance or fear of new technology.",
  "Leverages technology—Applies knowledge of technology to improve work processes and results (e.g., enhance productivity, efficiency, collaboration, quality, or customer satisfaction); uses technology to solve work-related problems, find new methods to enhance results, and create new business opportunities."
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


pObj = docx.createP();
pObj.addText('What strategies have you employed to ensure a major new directive from senior management was carried out?', { color: '000088', font_face: 'calibri', font_size: 12, indent:-300 })



var openingTable = [
  [{
    val: "Response",
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


var blankBox = [
  [{
    val: "\r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n ",
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



var scale = [
  [{
    val: "0",
    opts: {
      cellColWidth: 1625,
      color: '27348B',
      b:true,
      sz: '24',
      align: "center",
      vAlign: "center",
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      },

    }
  },
  {
    val: "1",
    opts: {
      b:true, // b = BOLD TEXT
      color: "27348B",
      align: "center",
      vAlign: "center",
      cellColWidth: 1625,
      sz: '24',
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "2",
    opts: {
      align: "center",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1625,
      b:true,
      sz: '24',
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "3",
    opts: {
      align: "center",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1625,
      b:true,
      sz: '24',
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "4",
    opts: {
      align: "center",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1625,
      b:true,
      sz: '24',
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      }
    }
  },
  {
    val: "5",
    opts: {
      align: "center",
      vAlign: "center",
      color: '27348B',
      cellColWidth: 1625,
      b:true,
      sz: '24',
      shd: {
        fill: "FFFFFF",

        "themeFillTint": "80"
      }
    }
  }
],

  [ { val:'Does Not Meet', opts: {sz:20}}, { val:'Limited', opts: {sz:20}}, { val:'Basic', opts: {sz:20}},
  { val:'Proficient', opts: {sz:20}},
{ val:'Excellent', opts: {sz:20}},{ val:'Outstanding', opts: {sz:20}}, ],



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


docx.putPageBreak();

// END OF PAGE 8


pObj = docx.createP()
pObj.addImage('uwalogo.jpg');


var closingTable = [
  [{
    val: "Closing",
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
  [{val:"What areas do you think you will need support or development in to effectively fulfill this role", opts:{sz:20}}]

]

var scaleStyle2 = {
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
  borderSize: 2, // To use this option, the 'borders' must set as true, default is 4
}

docx.createTable (closingTable, scaleStyle2);


var blankBox = [
  [{
    val: " \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n \r\n",
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






var questions1 = [
  [{
    val: "Do you have any questions for us?",
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


docx.createTable (questions1, scaleStyle3);



var blankBox = [
  [{
    val: " \r\n \r\n \r\n \r\n \r\n",
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

pObj = docx.createP(); // make space

var endInterview = [
  [{
    val: "End the Interview",
    opts: {
      cellColWidth: 9750,
      b:true,
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

var endStyle = {
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


docx.createTable (endInterview, endStyle);



pObj = docx.createP();

pObj.addText('Explain the next steps in selection process (i.e. UWA Talent team will communicate decision by...)',{ color: '27348B' })
pObj.addLineBreak()
pObj.addText('Thank the candidate for a productive interview',{ color: '27348B' })
pObj.addLineBreak()
pObj.addText('End the interview',{ color: '27348B' })













// Let's generate the Word document into a file:

let out = fs.createWriteStream('InterviewGuide.docx')

out.on('error', function(err) {
  console.log(err)
})

// Async call to generate the output file:
docx.generate(out)
