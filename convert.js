const { remote } = require("electron");
let filelocation = "";
const { dialog, Menu } = remote;

const fileLabel = document.getElementById("fileLabel");

const selectBtn = document.getElementById("selectBtn");
const ConvertBtn = document.getElementById("ConvertBtn");
const uploadSelectBtn = document.getElementById('uploadSelectBtn');
uploadSelectBtn.onclick = chooseUpload;

const nameInput = document.getElementById("nameInput");
const tableInput = document.getElementById("tableInput");


let converter

const template = [
  {
      label: 'Primke',
      click: async () => {
        converter = require('./primke')
        uploadLabel.childNodes[0].textContent="Primke"
      }
    },

  {
      label: 'Racuni',
      click: async () => {
          converter = require('./racuni');
          uploadLabel.childNodes[0].textContent="Računi"
      }
    },
    {
      label: 'Synesis Dnevnik',
      click: async () => {
          converter = require('./synesysis_dnevnik');
          uploadLabel.childNodes[0].textContent="Synesis Dnevnik"
      }
    },
    {
      label: 'Snesis URA IRA',
      click: async () => {
          converter = require('./synesysis_uraira');
          uploadLabel.childNodes[0].textContent="Snesis URA IRA"
      }
    }
]

async function chooseUpload() {
  const uploadOptions = Menu.buildFromTemplate(
      template
  );
  uploadOptions.popup();
}

selectBtn.onclick = async (e) => {
  file = await dialog.showOpenDialog({
    properties: ["openFile"],
    filters: [{ name: "xlsx only", extensions: ["xlsx","xls"] }],
  });
  filelocation = file.filePaths[0];
  fileLabel.childNodes[0].textContent= file.filePaths[0]
};

ConvertBtn.onclick = async (e) => {
  //console.log(filelocation)
  //console.log(converter)
  //console.log(nameInput.value)
    if(converter && filelocation != ""){
        //racuni.racuni(filelocation);
        converter.upload(filelocation, nameInput.value)
    }
    else{console.log("No file selected")}

};
