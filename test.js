
Office.initialize = function(reason){
    /*
    // functions
    function add42(num){
        write("hello from beyond");
        return num + 42000;
    }

    // definitions
    Excel.Script.CustomFunctions["CF"] = {};
    Excel.Script.CustomFunctions["CF"]["ADDTO44"] = {
        call: add42,
        description: "Returns the sum of a number and 42",
        helpUrl: "https://example.com/help.html",
        result: {
            resultType: Excel.CustomFunctionValueType.number,
            resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
        },
        parameters: [{
            name: "num",
            description: "The number be added",
            valueType: Excel.CustomFunctionValueType.number,
            valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
        },],
        options: {batch: false, stream: false,}
    };

    // registration
    Excel.run(function (context) {
        context.workbook.customFunctions.addAll();
        return context.sync();
    }).catch(function(error){
        write("Error:" + error);
    });
    */
    function write(myText){
        var fullHTML = "";
        Office.context.document.settings.refreshAsync(function(asyncResult){
            if(Office.context.document.settings.get("console")){
                fullHTML = Office.context.document.settings.get("console");
            }
            fullHTML += myText + "<br>";
            Office.context.document.settings.set("console",fullHTML);
            redrawConsole();
            Office.context.document.settings.saveAsync(function(asyncResult){});
        });
    }
    function redrawConsole(){
        if(Office.context.document.settings.get("console")){
            document.getElementById("debug").innerHTML = Office.context.document.settings.get("console");
        }
        else{
            document.getElementById("debug").innerHTML = "";
        }
    }
    function deleteConsole(){
        Office.context.document.settings.set("console","");
        redrawConsole();
        Office.context.document.settings.saveAsync(function(asyncResult){});
    }
    document.getElementById("redrawConsole").onclick = redrawConsole;
    document.getElementById("deleteConsole").onclick = deleteConsole;


    var debug = [];
    var debugUpdate = function(data){};
    
    /*
    function write(myText){
        debug.push([myText]);
        debugUpdate(debug);
        document.getElementById("debug").innerHTML += myText + "<br>";
    }
    */
    function myDebug(setResult){
        debugUpdate = setResult;
    }
/*
    Excel.Script.CustomFunctions["CFDEBUG"] = {};
    Excel.Script.CustomFunctions["CFDEBUG"]["DEBUG"] = {
        call: myDebug,
        description: "Outputs debug info to cells",
        helpUrl: "https://example.com/help.html",
        result: {
            resultType: Excel.CustomFunctionValueType.string,
            resultDimensionality: Excel.CustomFunctionDimensionality.matrix,
        },
        parameters: [
        ],
        options: {
            batch: false,
            stream: true,
        }
    };
*/
    var mycode;
    var hiddenPane = true;
    localStorage.setItem('readytoupdate', 0);

    if(!localStorage){
        write("no local storage");
    }
    else if(localStorage.getItem("mycode").length > 10){
        mycode = localStorage.getItem("mycode");
        eval(mycode);
        localStorage.setItem('readytoupdate', 0);
    }
    else{
        mycode = [
            '// functions',
            'function add42(num){',
            '\treturn num + 42;',
            '}',
            '',
            '// definitions',
            'Excel.Script.CustomFunctions["CF"] = {};',
            'Excel.Script.CustomFunctions["CF"]["ADDTO42"] = {',
            '\tcall: add42,',
            '\tdescription: "Returns the sum of a number and 42",',
            '\thelpUrl: "https://example.com/help.html",',
            '\tresult: {',
            '\t\tresultType: Excel.CustomFunctionValueType.number,',
            '\t\tresultDimensionality: Excel.CustomFunctionDimensionality.scalar,',
            '\t},',
            '\tparameters: [{',
            '\t\tname: "num",',
            '\t\tdescription: "The number be added",',
            '\t\tvalueType: Excel.CustomFunctionValueType.number,',
            '\t\tvalueDimensionality: Excel.CustomFunctionDimensionality.scalar,',
            '\t},],',
            '\toptions: {batch: false, stream: false,}',
            '};',
            '',
            '// registration',
            'Excel.run(function (context) {',
            '\tcontext.workbook.customFunctions.addAll();',
            '\treturn context.sync();',                
            '}).catch(function(error){',
            '\twrite("Error:" + error);',
            '});',
        ].join('\n');
    }
    
    var editor;

    require.config({ paths: { 'vs': 'monaco/min/vs' }});
    require(['vs/editor/editor.main'], function() {
        editor = monaco.editor.create(document.getElementById('container'), {
            value: mycode,
            codeLens: false,
            glyphMargin: false,
            language: 'javascript'
        });
    });

    
    
    document.getElementById("run").onclick = function(){runcode();};

    function runcode(){
        hiddenPane = false;
        mycode = editor.getValue();
        localStorage.setItem('mycode', mycode);
        localStorage.setItem('readytoupdate', 1);
        eval(mycode);
        document.getElementById("load").innerHTML = "Loading...";
        setTimeout(function(){
            document.getElementById("load").innerHTML = "Ready";
        },2000);
        write("ran code");
        
    }

    // start checking when we need to run
    pingForUpdate();
    
    //setInterval(messagePing, 3100);

    function pingForUpdate(){
        write("Hidden? " + hiddenPane + "; Readytoupdate? " + localStorage.getItem('readytoupdate'));
        
        if(!localStorage){
            write("no localstorage");
        }
        else{
            if(hiddenPane){
                
                if(localStorage.getItem('readytoupdate') == 1){
                    mycode = localStorage.getItem(mycode);
                    write("code changed " +localStorage.getItem('readytoupdate'));
                    eval(mycode);
                    localStorage.setItem('readytoupdate', 0);
                }
                setTimeout(pingForUpdate,2000);
            } 
        }

        
        
    }

    function messagePing(){
        write("am I hidden?" + hiddenPane);
        
    }
    
};