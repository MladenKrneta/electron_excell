const readXlsxFile = require('read-excel-file/node');
const { Client } = require("pg");
const moment = require('moment');
const aws = require('./config.json')
const { dialog } = remote;

const schema = {
    'Warehouse': {
      prop: 'warehouse',
      type: String
    },
    'Datum temelja': {
      prop: 'datumtemelja',
      type: Date,
    },
    'Oznaka temelja': {
        prop: 'oznakatemelja',
        type: String,
    },
    'OIB dobavljača': {
        prop: 'oib',
        type: String,
    },
    'Dobavljač': {
        prop: 'dobavljac',
        type: String,
    },
    'Način plaćanja': {
        prop: 'nacinplacanja',
        type: String,
    },
    'Broj računa': {
        prop: 'brojracuna',
        type: String,
    },
    'Broj dostavnice': {
        prop: 'brojdostavnice',
        type: String,
    },
    'Opis': {
        prop: 'opis',
        type: String,
    },
    'Ukupno primka': {
        prop: 'ukupnoprimka',
        type: Number,
    },
    'Poziv na broj': {
        prop: 'pozivnabroj',
        type: String,
    },
    'Datum dospijeća': {
        prop: 'datumdospijeca',
        type: Date,
    },
    'Roba-Osnovica 25%': {
        prop: 'robaosnovica25',
        type: Number,
    },
    'Roba-Osnovica 13%': {
        prop: 'robaosnovica13',
        type: Number,
    },
    'Roba-Osnovica 5%': {
        prop: 'robaosnovica5',
        type: Number,
    },
    'Roba-Porez 25%': {
        prop: 'robaporez25',
        type: Number,
    },
    'Roba-Porez 13%': {
        prop: 'robaporez13',
        type: Number,
    },
    'Roba-Porez 5%': {
        prop: 'robaporez5',
        type: Number,
    },
    'Roba-Povratna naknada': {
        prop: 'robapovratnanaknada',
        type: Number,
    },
    'Trošak-Osnovica 25%': {
        prop: 'trosakosnovica25',
        type: Number,
    },
    'Trošak-Osnovica 13%': {
        prop: 'trosakosnovica13',
        type: Number,
    },
    'Trošak-Osnovica 5%': {
        prop: 'trosakosnovica5',
        type: Number,
    },
    'Trošak-Porez 25%': {
        prop: 'trosakporez25',
        type: Number,
    },
    'Trošak-Porez 13%': {
        prop: 'trosakporez13',
        type: Number,
    },
    'Trošak-Porez 5%': {
        prop: 'trosakporez5',
        type: Number,
    },
    'Artikl': {
        prop: 'artikl',
        type: String,
    },
    'Nabavna cijena': {
        prop: 'nabavnacijena',
        type: Number,
    },
    'Količina': {
        prop: 'kolicina',
        type: Number,
    },
    'Iznos rabata': {
        prop: 'iznosrabata',
        type: Number,
    },
    'Ukupno': {
        prop: 'ukupno',
        type: Number,
    },
  }
  
  function ExcelDateToJSDate(date) {
    return new Date(Math.round((date - 25569)*86400*1000));
  }

  const upload = async (excel) =>{
    try{
        
        const client = new Client(aws);
        await client.connect();
    
        readXlsxFile(excel, { schema }).then((rows) => {
    
            t = "INSERT INTO sumarak.remaris_ulazni VALUES";
    
            let v = ''
            rows.rows.map((x,index)=>{
                //console.log(x)
                 v = v + "(" +
                "'" + x.warehouse + "'" + "," +
                "'" + moment(x.datumtemelja).format() + "'" + "," +
                "'" + x.oznakatemelja + "'" + "," +
                "'" + x.oib + "'" + "," +
                "'" + x.dobavljac + "'" + "," +
                "'" + x.nacinplacanja + "'" + "," +
                ( x.brojracuna? ("'"+x.brojracuna + "'" ) : ('null'))+"," +
                ( x.brojdostavnice? ("'"+x.brojdostavnice + "'" ) : ('null'))+"," +
                ( x.opis? ("'"+x.opis + "'" ) : ('null'))+"," +
                "'" + x.ukupnoprimka + "'" + "," +
                ( x.pozivnabroj? ("'"+x.pozivnabroj + "'" ) : ('null'))+"," +
                ( x.datumdospijeca? ("'"+moment(x.datumdospijeca).format() + "'" ) : ('null'))+"," +
                ( x.robaosnovica25? ("'"+x.robaosnovica25 + "'" ) : ('null'))+"," +
                ( x.robaosnovica13? ("'"+x.robaosnovica13 + "'" ) : ('null'))+"," +
                ( x.robaosnovica5? ("'"+x.robaosnovica5 + "'" ) : ('null'))+"," +
                ( x.robaporez25? ("'"+x.robaporez25 + "'" ) : ('null'))+"," +
                ( x.robaporez13? ("'"+x.robaporez13 + "'" ) : ('null'))+"," +
                ( x.robaporez5? ("'"+x.robaporez5 + "'" ) : ('null'))+"," +
                ( x.robapovratnanaknada? ("'"+x.robapovratnanaknada + "'" ) : ('null'))+"," +
                ( x.trosakosnovica25? ("'"+x.trosakosnovica25 + "'" ) : ('null'))+"," +
                ( x.trosakosnovica13? ("'"+x.trosakosnovica13 + "'" ) : ('null'))+"," +
                ( x.trosakosnovica5? ("'"+x.trosakosnovica5 + "'" ) : ('null'))+"," +
                ( x.trosakporez25? ("'"+x.trosakporez25 + "'" ) : ('null'))+"," +
                ( x.trosakporez13? ("'"+x.trosakporez13 + "'" ) : ('null'))+"," +
                ( x.trosakporez5? ("'"+x.trosakporez5 + "'" ) : ('null'))+"," +
                "'" + x.artikl + "'" + "," +
                "'" + x.nabavnacijena + "'" + "," +
                "'" + x.kolicina + "'" + "," +
                "'" + x.iznosrabata + "'" + "," +
                "'" + x.ukupno + "'" +  ")" 
                
                if(index + 1===rows.rows.length){
                    v = v+";" 
                }else{
                    v = v+","
                }
            })
            kveri = t + v
            // console.log(kveri)
            client.query(kveri)
            .then((res)=>{
              dialog.showMessageBox({ message: "Instert uspjesno napravljen!\r\nResponse: "+res.command + "\r\nCount: " + res.rowCount})
    
            })
            .catch((e)=>{
              console.log( e)
              dialog.showMessageBox({ message: 'An Error has occured! Read console for full details or Contact the creator :P'})
            })
        
        })
    
        } 
        catch (err) {
            console.log(err);
            // ... error checks
        }
}


exports.upload = upload;