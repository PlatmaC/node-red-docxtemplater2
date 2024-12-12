const axios = require('axios');
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

const convertTemplate = async ({ templateDocx, parameters, outputType, inputType }) => {
  let content;
  if (inputType==="nodebuffer"){
  content = templateDocx;
  } else if (inputType==="url"){
    content= await axios({
      method: 'get',
      url: templateDocx,
      responseType:"arraybuffer"
      });
      content = content.data;
  }
  content=content.toString('binary');

  const zip = new PizZip(content);
  const expressionParser = require('docxtemplater/expressions.js');
  const doc = new Docxtemplater(zip, {
    parser: expressionParser,
    paragraphLoop: true,
    linebreaks: true,
  });
  doc.render(parameters);
  const buf = doc.getZip().generate({
    type: outputType,
    compression: "DEFLATE",
  });
   return buf;
};

module.exports = function (RED) {
  function docxtemplateeditor(config) {
    RED.nodes.createNode(this, config);
    const node = this;

    node.on("input", async function (msg) {
      const templateDocx = config.templateDocx || msg.templateDocx;
      const inputType = config.inputType || msg.inputType;
      const outputType = config.outputType || msg.outputType;
      const parameters = msg.payload || {};
      try {
        const convertedTemplate = await convertTemplate({
          templateDocx,
          parameters,
          outputType,
          inputType
        });
        msg.payload = convertedTemplate;
        node.send(msg);
      } catch (error) {
        node.error(error);
      }
    });
  }
  RED.nodes.registerType("docxtemplateeditor", docxtemplateeditor);
};
