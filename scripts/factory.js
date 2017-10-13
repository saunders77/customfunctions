Office.initialize = function(reason){
    document.body.innerHTML += "It's an old world!!";
    function machinePurchaseDate(machineID) {
        
        return "7-24-2003";
    }
    document.body.innerHTML += "000";
    document.body.innerHTML += "customfunctions is " + Excel.Script.CustomFunctions;
    Excel.Script.CustomFunctions = {};
    Excel.Script.CustomFunction["CF"] = {};
    Excel.Script.CustomFunctions["CF"]["MACHINEPURCHASEDATE"] = {
        call: machinePurchaseDate,
        description: "Fetches the date when the machine was purchased",
        helpUrl: "https://michael-saunders.com/chocolate/help.html",
        result: {
            resultType: Excel.CustomFunctionValueType.string,
            resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
        },
        parameters: [
            {
                name: "Machine ID",
                description: "The ID code of the machine being queried",
                valueType: Excel.CustomFunctionValueType.string,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
        ],
        options: {
            batch: false,
            stream: false,
        }
    };

    document.body.innerHTML += "222";
    Excel.run(function (context) {
        context.workbook.customFunctions.addAll();
        return context.sync().then(function(){
            document.body.innerHTML += "returned from the sync";
        });
    
    }).catch(function(error){
        document.body.innerHTML += "error" + error;
    });
    document.body.innerHTML += "333";
};