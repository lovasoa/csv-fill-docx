#!/usr/bin/env node
/**
  @author Ophir LOJKINE
  @license BSD_3Clause
**/
var fs = require('fs');
var csv = require('csv-parser');
var DocxTemplater = require('docxtemplater');
var util = require('util');

function tplrep(s, dict) {
  for (i in dict) {
    s = s.replace('{' + i + '}', dict[i]);
  }
  return s;
}

function addZeroes(s, n) {
  while (s.length < n) s = '0' + s;
  return s;
}

function createFiller(outfiletpl) {
  //Load the docx file as a binary
  var content = fs.readFileSync(outfiletpl,"binary");
  var i = 0;

  return function fill(data) {
    var outfilename = tplrep(outfiletpl, data);
    if (outfilename === outfiletpl) outfilename = addZeroes(++i, 2) + " - " + outfilename;
    console.log("Generating '"+outfilename+"'...");
    var doc=new DocxTemplater(content);
    doc.setData(data);
    doc.render();
    var buf = doc.getZip().generate({type:"nodebuffer"});
    fs.writeFileSync(outfilename,buf);
  }
}

if (process.argv.length < 4) {
  console.log("Usage: ./csvfill.js input.csv template.docx");
  process.exit(1);
}
var ifname = process.argv[2];
var tplname = process.argv[3];

fs.createReadStream(ifname).pipe(csv({
  separator: ';'
})).on("data", createFiller(tplname));
