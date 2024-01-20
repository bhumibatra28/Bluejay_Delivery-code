const exceljs = require("exceljs"); // Import the exceljs library
const fs = require("fs");
const data = require("./Data.json");
async function readExcel() {
  const map = {}; //hashmap to store id with its corresponding time
  try {
    // 1. Load the Excel file:
    const workbook = new exceljs.Workbook();
    await workbook.xlsx.readFile("Assignment_Timecard.xlsx");

    // 2. Access the first worksheet:
    const worksheet = workbook.getWorksheet(1); // Index 1 represents the first sheet

    //reading the excel sheet and adding it to a json file
    // 3. Iterate through rows and cells:      
    // for (let row = 2; row <= worksheet.actualRowCount; row++) {
    //   const emprow = worksheet.getRow(row);
    //   const id = emprow.getCell(1).value;
    //   const timeIn = emprow.getCell(3).value;
    //   const timeOut = emprow.getCell(4).value;
    //   const name = emprow.getCell(8).value;
    //   if (map[id]) {
    //     map[id].timeIn.push(timeIn);
    //     map[id].timeOut.push(timeOut);
    //   } else {
    //     map[id] = {
    //       timeIn: [],
    //       timeOut: [],
    //       name: name,
    //     };
    //   }
    //   console.log(row + "row read");
    // }
    // fs.writeFile("Data.json", JSON.stringify(map), (e) => {
    //   if (e) console.log("error in saving file", e);
    // });
    // console.log("saved");
 

       //first part
    console.log('First part') 
    for (let key in data) {
      const set= new Set();      //creating set to find distinct dates
      for (let i in data[key].timeIn) {
        const dateTimeString = data[key].timeIn[i];   //timein date
        const date = new Date(dateTimeString);       //day of the time using inbuilt Date class
        set.add(date.getDate());
      }
      let count=0;
      const arr= Array.from(set);
      for(let i=0;i<arr.length-1;i++){    //checking for users dates
          if(arr[i+1]-arr[i]!=1)
          break;
          else
          count=count+1;
      }
      if(count>=7)
      console.log(data[key].name);
    }

      //second part 
      console.log('-----------')
      console.log('Second part')
      for (let key in data) {
      for(let i=0;i<data[key].timeIn.length-1;i++){
        const dateTimeIn = data[key].timeIn[i+1];
        const INtime= new Date(dateTimeIn);
        const In=INtime.getHours();
        const dateTimeout= data[key].timeOut[i];
        const Outtime= new Date(dateTimeout);
        const Out=Outtime.getHours();
        // console.log(In,Out);
        if(INtime.getDate()==Outtime.getDate())
        {
          if((In-Out)<10 && (In-Out)>1){
          console.log(data[key].name,In,Out);
          }
        }
      }
    }


      //third part
      console.log('-----------')
      console.log('Third part')
      for (let key in data) {
      let sum=0;
      let previous=0;
      for(let i=0;i<data[key].timeIn.length;i++){
          const dateTimeIn = data[key].timeIn[i];
          const INtime= new Date(dateTimeIn);
          const In=INtime.getHours();
          const dateTimeout= data[key].timeOut[i];
          const Outtime= new Date(dateTimeout);
          const Out=Outtime.getHours();
          if(previous===INtime.getDate()){
            sum+= Out-In;
          }
          else
          {
            if(sum>=14){
            console.log(data[key].name,sum);
            }
            sum= Out-In;
            previous= INtime.getDate();
          }

        }
    }

  } catch (error) {
    console.error("Error reading Excel file:", error);
  }
}

readExcel(); // Call the function to start reading
