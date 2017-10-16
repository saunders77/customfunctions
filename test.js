
Office.initialize = function(reason){

    var debug = [];
    var debugUpdate = function(data){};
    function write(myText){
        debug.push([myText]);
        debugUpdate(debug);
        document.getElementById("debug").innerHTML += myText + "<br>";
    }

    function myDebug(setResult){
        debugUpdate = setResult;
    }

    var mycode;
    var needToRegister = true;
    localStorage.setItem('readytoupdate', false);

    if(!localStorage){
        write("no local storage");
    }
    else if(localStorage.getItem("mycode").length > 10){
        mycode = localStorage.getItem("mycode");
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
        needToRegister = false;
        mycode = editor.getValue();
        localStorage.setItem('mycode', mycode);
        localStorage.setItem('readytoupdate', true);
        eval(mycode);
        document.getElementById("load").innerHTML = "Loading...";
        setTimeout(function(){
            document.getElementById("load").innerHTML = "Ready";
        },2000);
        write("ran code");
        
    }

    // start checking when we need to run
    pingForUpdate();
    

    function pingForUpdate(){
        if(!localStorage){
            write("no localstorage");
        }
        else{
            if(needToRegister){
                
                if(localStorage.getItem('readytoupdate') === true){
                    mycode = localStorage.getItem(mycode);
                    write("code changed " +localStorage.getItem('readytoupdate'));
                    eval(mycode);
                    localStorage.setItem('readytoupdate', false);
                }
                setTimeout(pingForUpdate,2000);
            } 
        }
        
    }
    
};