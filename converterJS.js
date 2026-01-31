// 'x' steht für Code, den wir aus Codebeispielen (alle im Quellenverzeichnis verlinkt) abgewandelt haben
//Verlinkung zum html Dokument
document.getElementById("demoA").onchange = evt => //x
{

  //liest die Exceldatei
  var reader = new FileReader(); //x

  
  //sorgt dafür, dass folgendes ausgeführt wird, während die Datei läd
  reader.addEventListener("loadend", evt => //x
  {

    //erstellt Tabelle; Verlinkung zur html Tabelle
    var table = document.getElementById("demoB"); //x
    //leere Tabelle
    table.innerHTML = ""; //x

    
    //liest Exceldatei ein und speichert sie als Objekt
    var workbook = XLSX.read(evt.target.result, {type: "array"}), //x
        //speichert Zelle A1 der Exceldatei
        worksheet = workbook.Sheets[workbook.SheetNames[0]], //x
        //speichert die beschriebenen Excelzellen
        range = XLSX.utils.decode_range(worksheet["!ref"]); //x



    //Längen- & Breitenkoordinaten vertauschen
    for (let row=range.s.r; row<=range.e.r; row++) 
    {

        var tausch = worksheet[XLSX.utils.encode_cell({r:row, c:4})];
        worksheet[XLSX.utils.encode_cell({r:row, c:4})] = worksheet[XLSX.utils.encode_cell({r:row, c:5})];
        worksheet[XLSX.utils.encode_cell({r:row, c:5})] = tausch;

    }

    

    //höchster Wert im Array Funktion
    function max(pArray)
    {
      let x = 0;
      for(let i = 0; i < pArray.length; i++)
      {
        if(x < pArray[i][2])
        {
          x = pArray[i][2];
        }
      }
      return x;
    }

    //höchster Wert im Array Funktion
    function min(pArray)
    {
      let x = 10000;
      for(let i = 0; i < pArray.length; i++)
      {
        if(x > pArray[i][2])
        {
          x = pArray[i][2];
        }
      }
      return x;
    }

    //Differenz der Werte
    function diff(pArray)
    {
      var  d = max(pArray) - min(pArray);
      return d;
    }
    
    

    //Zwischenspeicher Array erstellen
    const zwischenspeicher = [];
    zwischenspeicher.length = range.e.r;
    //Array zu 2D-Array machen, indem in jedem Feld ein Array gespeichert wird
    for(let i = 0; i < zwischenspeicher.length; i++)
    {
      zwischenspeicher[i] = [];
    }
    var z = 0;

    //Koordinaten im Zwischenspeicher Array speichern
    var inval = 0;
    for (let row=1; row<=range.e.r; row++) 
    {
      for (let col=range.s.c; col<=range.e.c; col++) 
      { 
        if(col == 4 || col == 5)
        {
          //Zugriff auf jeweilige Zelle aus dem Exceldokument   
          xcell = worksheet[XLSX.utils.encode_cell({r:row, c:col})]; //x
          //wandelt in String um, falls Zelle != null
          //speichert nichts, falls Zelle == null (try-catch für Nullpointerexception)
          let temp = xcell ? String(xcell.v): ""; //x
          //leere Zellen und invalide Werte werden rausgefiltert
          if(temp !== "Inval.")
          {
            if(temp !== "")
            {
              if(col==4)
              {
                //Nachkommastellen werden an der richtigen Stelle hinzugefügt
                temp = "[" + temp.slice(0,1) + "." + temp.slice(1) + ",";
                //Wert wird im Array gespeichert
                zwischenspeicher[z][0] = temp;
                
              }
              else if(col == 5)
              {
                temp = temp.slice(0,2) + "." + temp.slice(2) + "],";
                zwischenspeicher[z][1] = temp;
                  
              }
            }
          }
        }
      }
      z++;
    }


    //Komma löschen (wichtig für die Darstellung für die Google-Earth-API)
    var last = 0;
    for (let row=range.s.r; row< range.e.r; row++) 
    {
        if(zwischenspeicher[row][1] !== undefined)
        {
          last = row;
        }
    }
    var lastEX = worksheet[XLSX.utils.encode_cell({r:last, c:5})];
    var tempLast = lastEX ? String(lastEX.v): "";
    zwischenspeicher[last][1] = tempLast.slice(0,2) + "." + tempLast.slice(2) + "]];";





    //finales Array für die Koordinaten
    const speicher = [];
    speicher.length = range.e.r;
    for(let i = 0; i < speicher.length; i++)
    {
      speicher[i] = [];
    }

    //leere Arrayfelder am Anfang werden entfernt
    var sz = 0;
    for(let i = 0; i< speicher.length; i++)
    {
      if(zwischenspeicher[i][0] !== undefined)
      {
        speicher[sz][0] = zwischenspeicher[i][0];
        speicher[sz][1] = zwischenspeicher[i][1];
        sz++;
      }
    }



    //überflüssige Felder am Ende löschen
    var lastS = 0;
    for (let row=range.s.r; row< speicher.length; row++) 
    {
        if(speicher[row][1] !== undefined)
        {
          lastS = row;
        }
    }

    //neue Arraylänge speichern
    speicher.length = lastS+1;


    //Klammer am Anfang hinzufügen
    speicher[0][0] = "[" + speicher[0][0];

    table.insertRow().insertCell().innerHTML = "var koordinaten = ";


    //Array in der html Tabelle ausgeben
    for (let row=range.s.r; row< speicher.length; row++) 
    {
      //neue Zeile wird erstellt
      let r = table.insertRow(); //x

      for (let col=range.s.c; col<=1; col++) 
      {
        //neue Zelle in der Zeile wird erstellt
        let c = r.insertCell(); //x
          //Wert wird in der Zelle gespeichert
          c.innerHTML = speicher[row][col]; //x
        

      }
    }
    table.insertRow().insertCell().innerHTML = "";

    /*table.insertRow().insertCell().innerHTML = "blä";

    table.insertRow().insertCell().innerHTML = "nö";
     table.insertRow().insertCell().innerHTML = "nö";
      table.insertRow().insertCell().innerHTML = "nö";
       table.insertRow().insertCell().innerHTML = "nö";
        table.insertRow().insertCell().innerHTML = "nö";
         table.insertRow().insertCell().innerHTML = "nö";

    */















   //
   //sortiertes Array mit Höhe
   //Zwischenspeicher Array erstellen
    const zsHoehe = [];
    zsHoehe.length = range.e.r;
    for(let i = 0; i < zsHoehe.length; i++)
    {
      zsHoehe[i] = [];
    }
    z = 0;

    
    for (let row=1; row<=range.e.r; row++) 
    {
      for (let col=range.s.c; col<=range.e.c; col++) 
      {
        if(col == 4 || col == 5 || col == 7)
        {
          xcell = worksheet[XLSX.utils.encode_cell({r:row, c:col})];
          let temp = xcell ? String(xcell.v): "";

          if(temp !== "" && temp !== "Inval.")
          {
            
              if(col == 4)
              {
                temp = "[" + temp.slice(0,1) + "." + temp.slice(1) + ",";
                zsHoehe[z][0] = temp;
                
              }
              else if(col == 5)
              {
                temp = temp.slice(0,2) + "." + temp.slice(2) + "],";
                zsHoehe[z][1] = temp;
                  
              }
              else if(col == 7)
              {
                temp = temp.slice(0,temp.length) + "." + temp.slice(temp.length);
                //Umwandlung des Strings in einen Float (Dezimalzahl)
                temp = parseFloat(temp);
                zsHoehe[z][2] = temp;
              }
          }
        }
      }
      z++;
    }





    //noch nicht finales Array Höhe
    const spHoehe = [];
    spHoehe.length = range.e.r;
    for(let i = 0; i < spHoehe.length; i++)
    {
      spHoehe[i] = [];
    }

    sz = 0;
    for(let i = 0; i< spHoehe.length; i++)
    {
      if(zsHoehe[i][0] !== undefined)
      {
        spHoehe[sz][0] = zsHoehe[i][0];
        spHoehe[sz][1] = zsHoehe[i][1];
        spHoehe[sz][2] = zsHoehe[i][2];
        sz++;
      }
    }



    //überflüssige Felder löschen
    spHoehe.length = lastS+1;


    //Array ausgeben
    /*for (let row=range.s.r; row< spHoehe.length; row++) 
    {

      let r = table.insertRow();

      for (let col=range.s.c; col<=2; col++) 
      {

        let c = r.insertCell();
        
          c.innerHTML = spHoehe[row][col];
        

      }
    }

    table.insertRow().insertCell().innerHTML = "blä";
   */
  


    //Array in Unterarrays gliedern 
    /* das bestehende 2D-Array wird - nach der Höhe sortiert - in mehrere 2D-Arrays unterteilt;
    dieselben werden in einem weiteren Array gespeichert => 3D-Array*/
    const spHtemp = [];
    spHtemp.length = lastS+1;
    //3D-Array erstellen
    for(let i = 0; i < spHtemp.length; i++)
    {
      spHtemp[i] = [];
      for(let k = 0; k < spHtemp.length; k++)
      {
        spHtemp[i][k] = [];
      }
    }


    sz = 0;
    //Variablen für den jetzigen Wert und einen Vergleichswert
    var vorgaenger = 0;
    var jetzig = 0;
    for(let i = 0; i< spHoehe.length; i++)
    {
      //einsortieren in die jeweiligen Unterarrays
      //Verwendung der max-Funktion zur Bestimmung der Abschnitte
      if(spHoehe[i][2] < (max(spHoehe)/6))
      {
         jetzig = 1;
      }
      if(spHoehe[i][2] >=(max(spHoehe)/6) && spHoehe[i][2] <(max(spHoehe)/6*2))
      {
         jetzig = 2;
      }
      if(spHoehe[i][2] >=(max(spHoehe)/6*2) && spHoehe[i][2] <(max(spHoehe)/6*3))
      {
         jetzig = 3;
      }
      if(spHoehe[i][2] >=(max(spHoehe)/6*3) && spHoehe[i][2] <(max(spHoehe)/6*4))
      {
         jetzig = 4;
      }
      if(spHoehe[i][2] >=(max(spHoehe)/6*4) && spHoehe[i][2] <(max(spHoehe)/6*5))
      {
         jetzig = 5;
      }
      if(spHoehe[i][2] >=(max(spHoehe)/6*5) && spHoehe[i][2] <(max(spHoehe)))
      {
         jetzig = 6;
      }

      //neues Unterarray wird erstellt, wenn der jetzige Wert nicht in dieselbe Kategorie wie der Vorgängerwert gehört
      if(vorgaenger !== jetzig)
      {
        sz++;
      }
      //alle Arraywerte werden an der richtigen Stelle im 3D-Array gespeichert
      spHtemp[sz][i][0] = spHoehe[i][0];
      spHtemp[sz][i][1] = spHoehe[i][1];
      spHtemp[sz][i][2] = spHoehe[i][2]; 
      spHtemp[sz][i][3] = jetzig;
      vorgaenger = jetzig;
    }





    //finales Array für Höhe
    const spHfin = [];
    spHfin.length = spHtemp.length;
    for(let i = 0; i < spHfin.length; i++)
    {
      spHfin[i] = [];
      for(let k = 0; k < spHfin.length; k++)
      {
        spHfin[i][k] = [];
      }
    }


    var ssz = 0;
    for(let i = 1; i < spHfin.length; i++)
    {
      if(spHtemp[i][0] !== undefined)
      {
        sz = 0;
        for(let k = 0; k < spHfin[0].length; k++)
        {
          if(spHtemp[i][k][0] !== undefined)
          {
            spHfin[ssz][sz][0] = spHtemp[i][k][0];
            spHfin[ssz][sz][1] = spHtemp[i][k][1];
            spHfin[ssz][sz][2] = spHtemp[i][k][2];
            spHfin[ssz][sz][3] = spHtemp[i][k][3];
            sz++;
          }
        }
        ssz++;
      }
      
    }

    //überflüssige Felder am Anfang löschen
    var lastK = 0;
    var lastI = 0;
    for (let i= 0; i< spHfin.length; i++) 
    {
      if(spHfin[i][0][0] !== undefined)
      {
        lastI = i;
        lastK = 0;
        for(let k = 0; k < spHfin[i].length; k++)  
        {
          if(spHfin[i][k][0] !== undefined)
          {
            lastK = k;
          }
        }
        spHfin[i].length = lastK+1;
      }
    }
    spHfin.length = lastI+1;
    

    //Kommata hinzufügen
    for (let row=0; row< spHfin.length; row++) 
    {
      tempLast = spHfin[row][spHfin[row].length-1][1];
      spHfin[row][spHfin[row].length-1][1] = tempLast.slice(0,-1) + ",";
    }
    
    //Klammern hinzufügen
    for (let row=0; row< spHfin.length; row++) 
    {
      spHfin[row][0][0] = "[" + spHfin[row][0][0];
    }
    spHfin[0][0][2] = "[" + spHfin[0][0][2];
    spHfin[0][0][3] = "[" + spHfin[0][0][3];
    spHfin[spHfin.length-1][spHfin[0].length][2] =  spHfin[spHfin.length-1][spHfin[0].length][2] + "];";
    spHfin[spHfin.length-1][spHfin[0].length][3] =  spHfin[spHfin.length-1][spHfin[0].length][3] + "];";

    table.insertRow().insertCell().innerHTML = "";

    //3D-Array ausgeben (Stufen)
    table.insertRow().insertCell().innerHTML = "var heigth = ";
    let r = table.insertRow();

    for(let i = 0; i < spHfin.length; i++)
    {
      
      for (let row=range.s.r; row< spHfin[i].length; row++) 
      {
        
        
        if(i == spHfin.length-1 && row == spHfin[0].length)
        {
          r.insertCell().innerHTML = spHfin[i][row][3];
        }
        else
        {
          r.insertCell().innerHTML = spHfin[i][row][3] + ",";
        }
       
       
      }
      
    }
    table.insertRow().insertCell().innerHTML = "";

    //Stufeneinteilung ausgeben
    table.insertRow().insertCell().innerHTML = "var heigthDiff = [";
    for(let i = 0; i < 6; i++)
    {
      if(i !== 5)
      {
        table.insertRow().insertCell().innerHTML = "'" + parseInt(max(spHoehe)/6) * (i) + " - " + parseInt(max(spHoehe)/6) * (i+1) + "',"; 
      }
      else
      {
        table.insertRow().insertCell().innerHTML = "'" + parseInt(max(spHoehe)/6) * (i) + " - " + parseInt(max(spHoehe)/6) * (i+1) + "'"; 
      }
    }
    table.insertRow().insertCell().innerHTML = "];";
    table.insertRow().insertCell().innerHTML = "";

    /*
    table.insertRow().insertCell().innerHTML = "nö";
    table.insertRow().insertCell().innerHTML = "nö";
    table.insertRow().insertCell().innerHTML = "nö";
    table.insertRow().insertCell().innerHTML = "nö";
    table.insertRow().insertCell().innerHTML = "nö";
    */


    //table.insertRow().insertCell().innerHTML = diff(spHoehe);













   //
   //sortiertes Array mit Zeit
   //Zwischenspeicher Array erstellen
    const zsTime = [];
    zsTime.length = range.e.r;
    for(let i = 0; i < zsTime.length; i++)
    {
      zsTime[i] = [];
    }
    z = 0;

    
    for (let row=1; row<=range.e.r; row++) 
    {

      for (let col=range.s.c; col<=range.e.c; col++) 
      {
        
        if(col == 4 || col == 5 || col == 0)
        {

            xcell = worksheet[XLSX.utils.encode_cell({r:row, c:col})];

          let temp = xcell ? String(xcell.v): "";
          let ref = worksheet[XLSX.utils.encode_cell({r:row, c:4})];
          let reffe = ref ? String(ref.v): "";
          if(reffe !== "" && reffe !== "Inval.")
          {
            
              if(col == 4)
              {
                temp = "[" + temp.slice(0,1) + "." + temp.slice(1) + ",";
                zsTime[z][0] = temp;
                
              }
              else if(col == 5)
              {
                temp = temp.slice(0,2) + "." + temp.slice(2) + "],";
                zsTime[z][1] = temp;
                  
              }
              else if(col == 0)
              {
                temp = temp.slice(0,-3) + "." + temp.slice(-3);
                temp = parseFloat(temp);
                zsTime[z][2] = temp;
              }
            
          }
        }
      }
      z++;
    }





    //noch nicht finales Array Zeit
    const spTime = [];
    spTime.length = range.e.r;
    for(let i = 0; i < spTime.length; i++)
    {
      spTime[i] = [];
    }

    sz = 0;
    for(let i = 0; i< spTime.length; i++)
    {
      if(zsTime[i][0] !== undefined)
      {
        spTime[sz][0] = zsTime[i][0];
        spTime[sz][1] = zsTime[i][1];
        spTime[sz][2] = zsTime[i][2];
        sz++;
      }
    }


    //überflüssige Felder löschen
    spTime.length = lastS+1;




    //Array ausgeben
    /*for (let row=range.s.r; row< spTime.length; row++) 
    {

      let r = table.insertRow();

      for (let col=range.s.c; col<=2; col++) 
      {

        let c = r.insertCell();
        
          c.innerHTML = spTime[row][col];
        

      }
    }

    table.insertRow().insertCell().innerHTML = "blä";
    */


    //Array in Unterarrays gliedern 
    /* das bestehende 2D-Array wird - nach der Zeit sortiert - in mehrere 2D-Arrays unterteilt;
    dieselben werden in einem weiteren Array gespeichert */
    const spTtemp = [];
    spTtemp.length = lastS+1;
    for(let i = 0; i < spTtemp.length; i++)
    {
      spTtemp[i] = [];
      for(let k = 0; k < spTtemp.length; k++)
      {
        spTtemp[i][k] = [];
      }
    }

    sz = 0;
    vorgaenger = 0;
    jetzig = 0;
    for(let i = 0; i< spTime.length; i++) 
    {
      if(spTime[i][2] < ((diff(spTime)/6) + min(spTime)))
      {
         jetzig = 1;
      }
      if(spTime[i][2] >=((diff(spTime)/6)+ min(spTime)) && spTime[i][2] <((diff(spTime)/6*2)+ min(spTime)))
      {
         jetzig = 2;
      }
      if(spTime[i][2] >=((diff(spTime)/6*2)+ min(spTime)) && spTime[i][2] <((diff(spTime)/6*3)+ min(spTime)))
      {
         jetzig = 3;
      }
      if(spTime[i][2] >=((diff(spTime)/6*3)+ min(spTime)) && spTime[i][2] <((diff(spTime)/6*4)+ min(spTime)))
      {
         jetzig = 4;
      }
      if(spTime[i][2] >=((diff(spTime)/6*4)+ min(spTime)) && spTime[i][2] <((diff(spTime)/6*5)+ min(spTime)))
      {
         jetzig = 5;
      }
      if(spTime[i][2] >=((diff(spTime)/6*5)+ min(spTime)) && spTime[i][2] <((diff(spTime))+ min(spTime)))
      {
         jetzig = 6;
      }


      if(vorgaenger !== jetzig)
      {
        sz++;
      }
      spTtemp[sz][i][0] = spTime[i][0];
      spTtemp[sz][i][1] = spTime[i][1];
      spTtemp[sz][i][2] = spTime[i][2]; 
      spTtemp[sz][i][3] = jetzig;
      vorgaenger = jetzig;
    }


    //finales Array für Zeit
    const spTfin = [];
    spTfin.length = spTtemp.length;
    for(let i = 0; i < spTfin.length; i++)
    {
      spTfin[i] = [];
      for(let k = 0; k < spTfin.length; k++)
      {
        spTfin[i][k] = [];
      }
    }


    ssz = 0;
    for(let i = 1; i < spTfin.length; i++)
    {
      if(spTtemp[i][0] !== undefined)
      {
        sz = 0;
        for(let k = 0; k < spTfin[0].length; k++)
        {
          if(spTtemp[i][k][0] !== undefined)
          {
            spTfin[ssz][sz][0] = spTtemp[i][k][0];
            spTfin[ssz][sz][1] = spTtemp[i][k][1];
            spTfin[ssz][sz][2] = spTtemp[i][k][2];
            spTfin[ssz][sz][3] = spTtemp[i][k][3];
            sz++;
          }
        }
        ssz++;
      }
      
    }

    //überflüssige Felder löschen
    lastK = 0;
    lastI = 0;
    for (let i= 0; i< spTfin.length; i++) 
    {
      if(spTfin[i][0][0] !== undefined)
      {
        lastI = i;
        lastK = 0;
        for(let k = 0; k < spTfin[i].length; k++)  
        {
          if(spTfin[i][k][0] !== undefined)
          {
            lastK = k;
          }
        }
        spTfin[i].length = lastK+1;
      }
    }
    spTfin.length = lastI+1;
    

    //Kommata hinzufügen
    
    for (let row=0; row< spTfin.length; row++) 
    {
      tempLast = spTfin[row][spTfin[row].length-1][1];
      spTfin[row][spTfin[row].length-1][1] = tempLast.slice(0,-1) + ",";
    }
      
    
    //Klammern hinzufügen
    for (let row=0; row< spTfin.length; row++) 
    {
      spTfin[row][0][0] = "[" + spTfin[row][0][0];
    }
    spTfin[0][0][2] = "[" + spTfin[0][0][2];
    spTfin[0][0][3] = "[" + spTfin[0][0][3];
    spTfin[spTfin.length-1][spTfin[0].length][2] =  spTfin[spTfin.length-1][spTfin[0].length][2] + "];";
    spTfin[spTfin.length-1][spTfin[0].length][3] =  spTfin[spTfin.length-1][spTfin[0].length][3] + "];";

    table.insertRow().insertCell().innerHTML = "";

    //3D-Array ausgeben (Stufen)
    table.insertRow().insertCell().innerHTML = "var time = ";
    let q = table.insertRow();

    for(let i = 0; i < spTfin.length; i++)
    {
      
      for (let row=range.s.r; row< spTfin[i].length; row++) 
      {
        
        
        if(i == spTfin.length-1 && row == spTfin[0].length)
        {
          q.insertCell().innerHTML = spTfin[i][row][3];
        }
        else
        {
          q.insertCell().innerHTML = spTfin[i][row][3] + ",";
        }
       
       
      }
      
    }
    table.insertRow().insertCell().innerHTML = "";

    //Stufeneinteilung ausgeben
    table.insertRow().insertCell().innerHTML = "var timeDiff = [";
    for(let i = 0; i < 6; i++)
    {
      if(i !== 5)
      {
        table.insertRow().insertCell().innerHTML = "'" + (parseInt((diff(spTime)/6) * (i))) + " - " + (parseInt((diff(spTime)/6) * (i+1))) + "',"; 
      }
      else
      {
        table.insertRow().insertCell().innerHTML = "'" + (parseInt((diff(spTime)/6) * (i))) + " - " + (parseInt((diff(spTime)/6) * (i+1))) + "'"; 
      }
    }
    table.insertRow().insertCell().innerHTML = "];";
    table.insertRow().insertCell().innerHTML = "";


    /*
    table.insertRow().insertCell().innerHTML = "nö";
    table.insertRow().insertCell().innerHTML = "nö";
    table.insertRow().insertCell().innerHTML = "nö";
    table.insertRow().insertCell().innerHTML = "nö";
    table.insertRow().insertCell().innerHTML = "nö";
    table.insertRow().insertCell().innerHTML = "nö";
    */
    












   //
   //sortiertes Array mit Druck
   //Zwischenspeicher Array erstellen
    const zsPress = [];
    zsPress.length = range.e.r;
    for(let i = 0; i < zsPress.length; i++)
    {
      zsPress[i] = [];
    }
    z = 0;

    
    for (let row=1; row<=range.e.r; row++) 
    {

      //let r = table.insertRow();

      for (let col=range.s.c; col<=range.e.c; col++) 
      {
        
        if(col == 4 || col == 5 || col == 14)
        {
          //let c = r.insertCell(),

            xcell = worksheet[XLSX.utils.encode_cell({r:row, c:col})];

          let temp = xcell ? String(xcell.v): "";
          let ref = worksheet[XLSX.utils.encode_cell({r:row, c:4})];
          let reffe = ref ? String(ref.v): "";
          if(reffe !== "" && reffe !== "Inval.")
          {
            
              if(col == 4)
              {
                temp = "[" + temp.slice(0,1) + "." + temp.slice(1) + ",";
                zsPress[z][0] = temp;
                
              }
              else if(col == 5)
              {
                temp = temp.slice(0,2) + "." + temp.slice(2) + "],";
                zsPress[z][1] = temp;
                  
              }
              else if(col == 14)
              {
                temp = temp.slice(0,temp.length) + "." + temp.slice(temp.length);
                temp = parseFloat(temp);
                zsPress[z][2] = temp;
              }
            
            //c.innerHTML = temp;
          }
        }
      }
      z++;
    }





    //noch nicht finales Array Druck
    const spPress = [];
    spPress.length = range.e.r;
    for(let i = 0; i < spPress.length; i++)
    {
      spPress[i] = [];
    }

    sz = 0;
    for(let i = 0; i< spPress.length; i++)
    {
      if(zsPress[i][0] !== undefined)
      {
        spPress[sz][0] = zsPress[i][0];
        spPress[sz][1] = zsPress[i][1];
        spPress[sz][2] = zsPress[i][2];
        sz++;
      }
    }



    //überflüssige Felder löschen
    spPress.length = lastS+1;




    //Array ausgeben
    /*for (let row=range.s.r; row< spPress.length; row++) 
    {

      let r = table.insertRow();

      for (let col=range.s.c; col<=2; col++) 
      {

        let c = r.insertCell();
        
          c.innerHTML = spPress[row][col];
        

      }
    }

    table.insertRow().insertCell().innerHTML = "blä";
    */


    //Array in Unterarrays gliedern 
    /* das bestehende 2D-Array wird - nach dem Druck sortiert - in mehrere 2D-Arrays unterteilt;
    dieselben werden in einem weiteren Array gespeichert */
    const spPtemp = [];
    spPtemp.length = lastS+1;
    for(let i = 0; i < spPtemp.length; i++)
    {
      spPtemp[i] = [];
      for(let k = 0; k < spPtemp.length; k++)
      {
        spPtemp[i][k] = [];
      }
    }

    sz = 0;
    vorgaenger = 0;
    jetzig = 0;
    for(let i = 0; i< spPress.length; i++) 
    {
      if(spPress[i][2] < ((diff(spPress)/6)+ min(spPress)))
      {
         jetzig = 1;
      }
      if(spPress[i][2] >=((diff(spPress)/6)+ min(spPress)) && spPress[i][2] <((diff(spPress)/6*2)+ min(spPress)))
      {
         jetzig = 2;
      }
      if(spPress[i][2] >=((diff(spPress)/6*2)+ min(spPress)) && spPress[i][2] <((diff(spPress)/6*3)+ min(spPress)))
      {
         jetzig = 3;
      }
      if(spPress[i][2] >=((diff(spPress)/6*3)+ min(spPress)) && spPress[i][2] <((diff(spPress)/6*4)+ min(spPress)))
      {
         jetzig = 4;
      }
      if(spPress[i][2] >=((diff(spPress)/6*4)+ min(spPress)) && spPress[i][2] <((diff(spPress)/6*5)+ min(spPress)))
      {
         jetzig = 5;
      }
      if(spPress[i][2] >=((diff(spPress)/6*5)+ min(spPress)) && spPress[i][2] <((diff(spPress))+ min(spPress)))
      {
         jetzig = 6;
      }


      if(vorgaenger !== jetzig)
      {
        sz++;
      }
      spPtemp[sz][i][0] = spPress[i][0];
      spPtemp[sz][i][1] = spPress[i][1];
      spPtemp[sz][i][2] = spPress[i][2]; 
      spPtemp[sz][i][3] = jetzig;
      vorgaenger = jetzig;
    }


    //finales Array für Druck
    const spPfin = [];
    spPfin.length = spPtemp.length;
    for(let i = 0; i < spPfin.length; i++)
    {
      spPfin[i] = [];
      for(let k = 0; k < spPfin.length; k++)
      {
        spPfin[i][k] = [];
      }
    }


    ssz = 0;
    for(let i = 1; i < spPfin.length; i++)
    {
      if(spPtemp[i][0] !== undefined)
      {
        sz = 0;
        for(let k = 0; k < spPfin[0].length; k++)
        {
          if(spPtemp[i][k][0] !== undefined)
          {
            spPfin[ssz][sz][0] = spPtemp[i][k][0];
            spPfin[ssz][sz][1] = spPtemp[i][k][1];
            spPfin[ssz][sz][2] = spPtemp[i][k][2];
            spPfin[ssz][sz][3] = spPtemp[i][k][3];
            sz++;
          }
        }
        ssz++;
      }
      
    }

    //überflüssige Felder löschen
    lastK = 0;
    lastI = 0;
    for (let i= 0; i< spPfin.length; i++) 
    {
      if(spPfin[i][0][0] !== undefined)
      {
        lastI = i;
        lastK = 0;
        for(let k = 0; k < spPfin[i].length; k++)  
        {
          if(spPfin[i][k][0] !== undefined)
          {
            lastK = k;
          }
        }
        spPfin[i].length = lastK+1;
      }
    }
    spPfin.length = lastI+1;
    

  
  //Kommata hinzufügen
    for (let row=0; row< spPfin.length; row++) 
    {
      tempLast = spPfin[row][spPfin[row].length-1][1];
      spPfin[row][spPfin[row].length-1][1] = tempLast.slice(0,-1) + ",";
    }
      
    
    //Klammern hinzufügen
    for (let row=0; row< spPfin.length; row++) 
    {
      spPfin[row][0][0] = "[" + spPfin[row][0][0];
    }
    spPfin[0][0][2] = "[" + spPfin[0][0][2];
    spPfin[0][0][3] = "[" + spPfin[0][0][3];
    spPfin[spPfin.length-1][spPfin[0].length][2] =  spPfin[spPfin.length-1][spPfin[0].length][2] + "];";
    spPfin[spPfin.length-1][spPfin[0].length][3] =  spPfin[spPfin.length-1][spPfin[0].length][3] + "];";

    table.insertRow().insertCell().innerHTML = "";

    //3D-Array ausgeben (Stufen)
    table.insertRow().insertCell().innerHTML = "var pressure = ";
    var t = table.insertRow();

    for(let i = 0; i < spPfin.length; i++)
    {
      
      for (let row=range.s.r; row< spPfin[i].length; row++) 
      {
        
        
        if(i == spPfin.length-1 && row == spPfin[0].length)
        {
          t.insertCell().innerHTML = spPfin[i][row][3];
        }
        else
        {
          t.insertCell().innerHTML = spPfin[i][row][3] + ",";
        }
       
       
      }
      
    }
    table.insertRow().insertCell().innerHTML = "";

    //Stufeneinteilung ausgeben
    table.insertRow().insertCell().innerHTML = "var pressDiff = [";
    for(let i = 0; i < 6; i++)
    {
      if(i !== 5)
      {
        table.insertRow().insertCell().innerHTML = "'" + (parseInt((diff(spPress)/6) * (i)+ min(spPress))) + " - " + (parseInt((diff(spPress)/6) * (i+1)+ min(spPress))) + "',"; 
      }
      else
      {
        table.insertRow().insertCell().innerHTML = "'" + (parseInt((diff(spPress)/6) * (i)+ min(spPress))) + " - " + (parseInt((diff(spPress)/6) * (i+1)+ min(spPress))) + "'"; 
      }
    }
    table.insertRow().insertCell().innerHTML = "];";
    table.insertRow().insertCell().innerHTML = "";



    
    /*
    table.insertRow().insertCell().innerHTML = "nö";
    table.insertRow().insertCell().innerHTML = "nö";
    table.insertRow().insertCell().innerHTML = "nö";
    table.insertRow().insertCell().innerHTML = "nö";
    table.insertRow().insertCell().innerHTML = "nö";
    table.insertRow().insertCell().innerHTML = "nö";
    */
  
  
  
  
  
  
  
  
  
  
  
  
  
  
   //
   //sortiertes Array mit Luftfeuchtigkeit
   //Zwischenspeicher Array erstellen
   
    const zsHum = [];
    zsHum.length = range.e.r;
    for(let i = 0; i < zsHum.length; i++)
    {
      zsHum[i] = [];
    }
    z = 0;

    
    for (let row=1; row<=range.e.r; row++) 
    {
      for (let col=range.s.c; col<=range.e.c; col++) 
      {
        
        if(col == 4 || col == 5 || col == 15)
        {
            xcell = worksheet[XLSX.utils.encode_cell({r:row, c:col})];

          let temp = xcell ? String(xcell.v): "";
          let ref = worksheet[XLSX.utils.encode_cell({r:row, c:5})];
          let reffe = ref ? String(ref.v): "";
          if(reffe !== "" && reffe !== "Inval.")
          {
            
              if(col == 4)
              {
                temp = "[" + temp.slice(0,1) + "." + temp.slice(1) + ",";
                zsHum[z][0] = temp;
                
              }
              else if(col == 5)
              {
                temp = temp.slice(0,2) + "." + temp.slice(2) + "],";
                zsHum[z][1] = temp;
                  
              }
              else if(col == 15)
              {
                temp = temp.slice(0,temp.length-1) + "." + temp.slice(temp.length-1);
                temp = parseFloat(temp);
                zsHum[z][2] = temp;
              }
          }
        }
      }
      z++;
    }
    




    //noch nicht finales Array Luftfeuchtigkit
    
    const spHum = [];
    spHum.length = range.e.r;
    for(let i = 0; i < spHum.length; i++)
    {
      spHum[i] = [];
    }

    sz = 0;
    for(let i = 0; i< spHum.length; i++)
    {
      if(zsHum[i][0] !== undefined)
      {
        spHum[sz][0] = zsHum[i][0];
        spHum[sz][1] = zsHum[i][1];
        spHum[sz][2] = zsHum[i][2];
        sz++;
      }
    }
    


    //überflüssige Felder löschen
    spHum.length = lastS+1;
    


    
    //Array ausgeben
    /*for (let row=range.s.r; row< spHum.length; row++) 
    {

      let r = table.insertRow();

      for (let col=range.s.c; col<=2; col++) 
      {

        let c = r.insertCell();
        
          c.innerHTML = spHum[row][col];
        

      }
    }

    table.insertRow().insertCell().innerHTML = "blä";
    


    //Array in Unterarrays gliedern 
    /* das bestehende 2D-Array wird - nach der Luftfeuchtigkeit sortiert - in mehrere 2D-Arrays unterteilt;
    dieselben werden in einem weiteren Array gespeichert */
    /*
    const spHutemp = [];
    spHutemp.length = lastS+1;
    for(let i = 0; i < spHutemp.length; i++)
    {
      spHutemp[i] = [];
      for(let k = 0; k < spHutemp.length; k++)
      {
        spHutemp[i][k] = [];
      }
    }

    sz = 0;
    vorgaenger = 0;
    jetzig = 0;
    for(let i = 0; i< spHum.length; i++) 
    {
      if(spHum[i][2] < ((diff(spHum)/6) + min(spHum)))
      {
         jetzig = 1;
      }
      if(spHum[i][2] >=((diff(spHum)/6)+ min(spHum)) && spHum[i][2] < ((diff(spHum)/6*2)+ min(spHum)))
      {
         jetzig = 2;
      }
      if(spHum[i][2] >=((diff(spHum)/6*2)+ min(spHum)) && spHum[i][2] <((diff(spHum)/6*3)+ min(spHum)))
      {
         jetzig = 3;
      }
      if(spHum[i][2] >=((diff(spHum)/6*3)+ min(spHum)) && spHum[i][2] <((diff(spHum)/6*4)+ min(spHum)))
      {
         jetzig = 4;
      }
      if(spHum[i][2] >=((diff(spHum)/6*4)+ min(spHum)) && spHum[i][2] <((diff(spHum)/6*5)+ min(spHum)))
      {
         jetzig = 5;
      }
      if(spHum[i][2] >=((diff(spHum)/6*5)+ min(spHum)) && spHum[i][2] <100)
      {
         jetzig = 6;
      }


      if(vorgaenger !== jetzig)
      {
        sz++;
      }
      spHutemp[sz][i][0] = spHum[i][0];
      spHutemp[sz][i][1] = spHum[i][1];
      spHutemp[sz][i][2] = spHum[i][2]; 
      spHutemp[sz][i][3] = jetzig;
      vorgaenger = jetzig;
    }


    //finales Array für Luftfeuchtigkeit
    const spHufin = [];
    spHufin.length = spHutemp.length;
    for(let i = 0; i < spHufin.length; i++)
    {
      spHufin[i] = [];
      for(let k = 0; k < spHufin.length; k++)
      {
        spHufin[i][k] = [];
      }
    }


    ssz = 0;
    for(let i = 1; i < spHufin.length; i++)
    {
      if(spHutemp[i][0] !== undefined)
      {
        sz = 0;
        for(let k = 0; k < spHufin[0].length; k++)
        {
          if(spHutemp[i][k][0] !== undefined)
          {
            spHufin[ssz][sz][0] = spHutemp[i][k][0];
            spHufin[ssz][sz][1] = spHutemp[i][k][1];
            spHufin[ssz][sz][2] = spHutemp[i][k][2];
            spHufin[ssz][sz][3] = spHutemp[i][k][3];
            sz++;
          }
        }
        ssz++;
      }
      
    }

    //überflüssige Felder löschen
    lastK = 0;
    lastI = 0;
    for (let i= 0; i< spHufin.length; i++) 
    {
      if(spHufin[i][0][0] !== undefined)
      {
        lastI = i;
        lastK = 0;
        for(let k = 0; k < spHufin[i].length; k++)  
        {
          if(spHufin[i][k][0] !== undefined)
          {
            lastK = k;
          }
        }
        spHufin[i].length = lastK+1;
      }
    }
    spHufin.length = lastI+1;
    



    //Kommata hinzufügen
    for (let row=0; row< spHufin.length; row++) 
    {
      tempLast = spHufin[row][spHufin[row].length-1][1];
      spHufin[row][spHufin[row].length-1][1] = tempLast.slice(0,-1) + ",";
    }
    
    //Klammern hinzufügen
    for (let row=0; row< spHufin.length; row++) 
    {
      spHufin[row][0][0] = "[" + spHufin[row][0][0];
    }
    spHufin[0][0][2] = "[" + spHufin[0][0][2];
    spHufin[0][0][3] = "[" + spHufin[0][0][3];
    spHufin[spHufin.length-1][spHufin[0].length][2] =  spHufin[spHufin.length-1][spHufin[0].length][2] + "];";
    spHufin[spHufin.length-1][spHufin[0].length][3] =  spHufin[spHufin.length-1][spHufin[0].length][3] + "];";
    
    table.insertRow().insertCell().innerHTML = "";

    //3D-Array ausgeben (Stufen)
    table.insertRow().insertCell().innerHTML = "var humidity = ";
    var u = table.insertRow();

    for(let i = 0; i < spHufin.length; i++)
    {
      
      for (let row=range.s.r; row< spHufin[i].length; row++) 
      {
        
        
        if(i == spHufin.length-1 && row == spHufin[0].length)
        {
          u.insertCell().innerHTML = spHufin[i][row][3];
        }
        else
        {
          u.insertCell().innerHTML = spHufin[i][row][3] + ",";
        }
       
       
      }
      
    }
    table.insertRow().insertCell().innerHTML = "";

    //Stufeneinteilung ausgeben
    table.insertRow().insertCell().innerHTML = "var humDiff = [";
    for(let i = 0; i < 6; i++)
    {
      if(i == 5)
      {
        table.insertRow().insertCell().innerHTML = "'" + (parseInt((diff(spHum)/6) * (i)+ min(spHum))) + " - 100'"; 
      }
      else if(i == 0)
      {
        table.insertRow().insertCell().innerHTML = "'0 - " + (parseInt((diff(spHum)/6) * (i+1) + min(spHum))) + "',"; 
      }
      else
      {
        table.insertRow().insertCell().innerHTML = "'" + (parseInt((diff(spHum)/6) * (i)+ min(spHum))) + " - " + (parseInt((diff(spHum)/6) * (i+1) + min(spHum))) + "',";  
      }
    }
    table.insertRow().insertCell().innerHTML = "];";
    table.insertRow().insertCell().innerHTML = "";

    */


    /*
    table.insertRow().insertCell().innerHTML = "nö";
    table.insertRow().insertCell().innerHTML = "nö";
    table.insertRow().insertCell().innerHTML = "nö";
    table.insertRow().insertCell().innerHTML = "nö";
    table.insertRow().insertCell().innerHTML = "nö";
    table.insertRow().insertCell().innerHTML = "nö";
    */
  








  




//
//Graphen
//Arrays erstellen
const time = [];
time.length = spTime.length;
const heigth = [];
heigth.length = spHoehe.length;
const heigthSort = [];
heigthSort.length = spHoehe.length;
const pressure = [];
pressure.length = spPress.length;
const pressureSort = [];
pressureSort.length = spPress.length;
const humidity = [];
humidity.length = spHum.length;
const humiditySort = [];
humiditySort.length = spHum.length;



//Arrays nur mit der jeweiligen Excelspalte füllen
//(die bisherigen Arrays enthielten immer auch weitere Daten wie die Koordinaten)
for(let i = 0; i < spTime.length; i++)
{
     time[i] = spTime[i][2];
}
for(let i = 0; i < spHoehe.length; i++)
{
     heigth[i] = spHoehe[i][2];
}
for(let i = 0; i < spHoehe.length; i++)
{
     heigthSort[i] = spHoehe[i][2];
}
for(let i = 0; i < spPress.length; i++)
{
     pressure[i] = spPress[i][2];
}
for(let i = 0; i < spPress.length; i++)
{
     pressureSort[i] = spPress[i][2];
}
for(let i = 0; i < spHum.length; i++)
{
     humidity[i] = spHum[i][2];
}
for(let i = 0; i < spHum.length; i++)
{
     humiditySort[i] = spHum[i][2];
}



//Bubblesort nach der Höhe für weitere Graphen
for(let i = 0; i < heigthSort.length; i++)
{
  for(let k = 0; k < heigthSort.length-1; k++)
  {
    if(heigthSort[k] > heigthSort[k+1])
    {
      let x = heigthSort[k];
      let y = humiditySort[k];
      let a = pressureSort[k];
      heigthSort[k] = heigthSort[k+1];
      humiditySort[k] = humiditySort[k+1];
      pressureSort[k] = pressureSort[k+1];
      heigthSort[k+1] = x;
      humiditySort[k+1] = y;
      pressureSort[k+1] = a;
    }
  }
}


//einheitliche x-Achse für Höhe erstellen
const heigthAchse = []; 
let zl = 0;
heigthAchse.length = heigth.length;
for(let i = 0; i < heigthAchse.length; i++)
{
  
  heigthAchse[i] = (max(heigthSort)/heigthSort.length)*(i);
  table.insertRow().insertCell().innerHTML = heigthAchse[i];
}


  
//Verbindung zum html Dokument
const chart1 = new Chart('myChart1',
  {
    //Art des Graphen
    type: 'line',
    data: 
    {
      //Daten für die x-Achse
      labels: time,
      
      datasets: 
      [
        {
          label: 'Hoehenprofil',
          borderColor: '#900f',
          //Daten für die y-Achse
          data: heigth,
          
        }
      ]
    },
    //Beschriftung der Achsen
    options: 
    {
    scales: 
    {
      x: 
      {
        title:
        {
          display: true,
          text: 'Zeit (s)',
        }
      },
      y:
      {
        title: 
        {
          display: true,
          text: 'Hoehe (m)',
        }, 
      beginAtZero: true
      }
    }
   }
  }
)
const chart2 = new Chart('myChart2',
  {
    type: 'line',
    data: 
    {
      labels: time,
      datasets: 
      [
        {
          label: 'Druckschwankungen',
          borderColor: '#666f',
          data: pressure,
        }
      ]
    },
    options: 
    {
    scales: 
    {
      x: 
      {
        title:
        {
          display: true,
          text: 'Zeit (s)',
        }
      },
      y:
      {
        title: 
        {
          display: true,
          text: 'Druck (hPa)',
        }, 
      }
    }
   }
  }
)
const chart3 = new Chart('myChart3',
  {
    type: 'line',
    data: 
    {
      labels: heigthAchse,
      datasets: 
      [
        {
          label: 'Druck-Hoehen-Abhaengigkeit',
          borderColor: '#333f',
          data: pressureSort,
        }
      ]
    },
    options: 
    {
    scales: 
    {
      x: 
      {
        title:
        {
          display: true,
          text: 'Hoehe (m)',
        }
      },
      y:
      {
        title: 
        {
          display: true,
          text: 'Druck (hPa)',
        }, 
      }
    }
   }
  }
)
const chart4 = new Chart('myChart4',
  {
    type: 'line',
    data: 
    {
      labels: time,
      datasets: 
      [
        {
          label: 'Luftfeuchtigkeitsschwankungen',
          borderColor: '#0f0b',
          data: humidity,
        }
      ]
    },
    options: 
    {
    scales: 
    {
      x: 
      {
        title:
        {
          display: true,
          text: 'Zeit (s)',
        }
      },
      y:
      {
        title: 
        {
          display: true,
          text: 'Luftfeuchtigkeit (%)',
        }, 
      }
    }
   }
  }
)
const chart5 = new Chart('myChart5',
  {
    type: 'line',
    data: 
    {
      labels: heigthAchse,
      datasets: 
      [
        {
          label: 'Luftfeuchtigkeit-Hoehen-Abhaengigkeit',
          borderColor: '#080f',
          data: humiditySort,
        }
      ]
    },
    options: 
    {
    scales: 
    {
      x: 
      {
        title:
        {
          display: true,
          text: 'Hoehe (m)',
        }
      },
      y:
      {
        title: 
        {
          display: true,
          text: 'Luftfeuchtigkeit (%)',
        }, 
      }
    }
   }
  }
)


 });

  
  //liest die Exceldatei als ArrayBuffer Version ein
  reader.readAsArrayBuffer(evt.target.files[0]); //x

  

};
 

