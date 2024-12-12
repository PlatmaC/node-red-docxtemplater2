const axios  = require( 'axios');
const libre = require('libreoffice-convert');
const { promisify }  = require( 'util');
libre.convertAsync = promisify(libre.convert);


const convertTemplate = async ({ docx, outputType, inputType }) => {
  let content;
  if (inputType==="nodebuffer"){
  content = docx;
  } else if (inputType==="url"){
    content= await axios({
      method: 'get',
      url: docx,
      responseType:"arraybuffer"
      });
      content = content.data;
  }

  let res = await libre.convertAsync(content, "pdf", '')
  
  if (outputType==="base64"){
   res = Buffer.from(res).toString('base64');
    }

  return res;
};

module.exports = function (RED) {
  function docxtopdf(config) {
    RED.nodes.createNode(this, config);
    const node = this;

    node.on("input", async function (msg) {
      const docx = config.docx || msg.docx;
      const inputType = config.inputType || msg.inputType;
      const outputType = config.outputType || msg.outputType;
      try {
        const convertedDocx = await convertTemplate({
          docx,
          outputType,
          inputType
        });
        msg.payload = convertedDocx;
        node.send(msg);
      } catch (error) {
        node.error(error);
      }
    });
  }
  RED.nodes.registerType("docxtopdf", docxtopdf);
};
