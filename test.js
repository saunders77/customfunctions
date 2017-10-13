Office.initialize = function(reason){
    var debug = [];
    var debugUpdate = function(data){};
    function write(myText){
        debug.push([myText]);
        debugUpdate(debug);
    }

    function myDebug(setResult){
        debugUpdate = setResult;
    }
    
    // helper method for promises
    Excel.Promise = function (setResultFunction){
        return new OfficeExtension.Promise(function(resolve, reject){
            setResultFunction(resolve, reject);
        });
    }  
    
    // helper code for getting temperature
    var temps = {};
    temps["boiler"] = 104.3;
    temps["mixer"] = 44.0;
    temps["furnace"] = 586.9;
    furnaceHistory = [];
    function startTime(){
        temps["boiler"] += Math.pow(Math.random() - 0.45, 3) * 2;
        temps["mixer"] += Math.pow(Math.random() - 0.55, 3) * 2;
        temps["furnace"] += Math.pow(Math.random() - 0.40, 3) * 2;
        furnaceHistory.push([temps["furnace"]]);
        if(furnaceHistory.length > 50){
            furnaceHistory.shift();
        }
        setTimeout(startTime, 500);
    }
    startTime();
    function getTempFromServer(thermometerID, callback){
        setTimeout(function(){
            var data = {};
            data.temperature = temps[thermometerID].toFixed(1);
            callback(data);
        }, 200);
    }

    // demo functions

    function addTo42(num){
        return Excel.Promise(function(setResult, setError){
            setTimeout(function(){
                setResult(num + 42);
            }, 1000);
        });
    }
    
    function addTo42Fast(num) {
        return num + 42;
    }

    function getTemperature(thermometerID){ 
        return Excel.Promise(function(setResult, setError){ 
            getTempFromServer(thermometerID, function(data){ 
                setResult(data.temperature); 
            }); 
        }); 
    }

    function streamTemperature(thermometerID, interval, setResult){     
        if(thermometerID == "furnace"){
            temps["furnace"] = 630.2;
        }
        function getNextTemperature(){ 
            getTempFromServer(thermometerID, function(data){ 
                setResult(data.temperature); 
            }); 
            setTimeout(getNextTemperature, interval); 
        } 
        getNextTemperature(); 
    } 

    function secondHighestTemp(temperatures){ 
        var highest = -273, secondHighest = -273;
        for(var i = 0; i < temperatures.length;i++){
            for(var j = 0; j < temperatures[i].length;j++){
                if(temperatures[i][j] >= highest){
                    secondHighest = highest;
                    highest = temperatures[i][j];
                }
                else if(temperatures[i][j] >= secondHighest){
                    secondHighest = temperatures[i][j];
                }
            }
        }
        return secondHighest;
    }

    function trackTemperature(thermometerID, setResult){
        var output = [];
        
        for(var i = 0; i < 50; i++) output.push([0]);  
        if(thermometerID == "furnace"){
            output = furnaceHistory;
        } 
        function recordNextTemperature(){
            getTempFromServer(thermometerID, function(data){
                output.push([data.temperature]);
                output.shift();
                setResult(output);
            });
            setTimeout(recordNextTemperature, 500);
        }
        recordNextTemperature();
    } 

    
    Excel.Script.CustomFunctions = {};
    Excel.Script.CustomFunctions["MY"] = {};
    Excel.Script.CustomFunctions["MY"]["ADDTO42"] = {
        call: addTo42,
        description: "Returns the sum of a number and 42",
        helpUrl: "https://example.com/help.html",
        result: {
            resultType: Excel.CustomFunctionValueType.number,
            resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
        },
        parameters: [
            {
                name: "num",
                description: "The number be added",
                valueType: Excel.CustomFunctionValueType.number,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
        ],
        options: {
            batch: false,
            stream: false,
        }
    };
    Excel.Script.CustomFunctions["MY"]["DEBUG"] = {
        call: myDebug,
        description: "Returns debugging text",
        helpUrl: "https://example.com/help.html",
        result: {
            resultType: Excel.CustomFunctionValueType.string,
            resultDimensionality: Excel.CustomFunctionDimensionality.matrix,
        },
        parameters: [],
        options: {
            batch: false,
            stream: true,
        }
    };
    Excel.Script.CustomFunctions["FN"] = {};
    Excel.Script.CustomFunctions["FN"]["ADDTO42"] = {
        call: addTo42Fast,
        description: "Returns the sum of a number and 42, fast",
        helpUrl: "https://example.com/help.html",
        result: {
            resultType: Excel.CustomFunctionValueType.number,
            resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
        },
        parameters: [
            {
                name: "num",
                description: "The number be added",
                valueType: Excel.CustomFunctionValueType.number,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
        ],
        options: {
            batch: false,
            stream: false,
        }
    };
    Excel.Script.CustomFunctions["FN"]["GETTEMPERATURE"] = {
        call: getTemperature,
        description: "Returns the temperature of the boiler, mixer, or furnace",
        helpUrl: "https://example.com/help.html",
        result: {
            resultType: Excel.CustomFunctionValueType.number,
            resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
        },
        parameters: [
            {
                name: "thermometer ID",
                description: "The thermometer to be measured",
                valueType: Excel.CustomFunctionValueType.string,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
        ],
        options: {
            batch: false,
            stream: false,
        }
    };
    Excel.Script.CustomFunctions["FN"]["STREAMTEMPERATURE"] = {
        call: streamTemperature,
        description: "Streams the temperature of the boiler, mixer, or furnace",
        helpUrl: "https://example.com/help.html",
        result: {
            resultType: Excel.CustomFunctionValueType.number,
            resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
        },
        parameters: [
            {
                name: "thermometer ID",
                description: "The thermometer to be measured",
                valueType: Excel.CustomFunctionValueType.string,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
            {
                name: "interval",
                description: "The time between updates",
                valueType: Excel.CustomFunctionValueType.number,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
        ],
        options: {
            batch: false,
            stream: true,
        }
    };
    Excel.Script.CustomFunctions["FN"]["SECONDHIGHESTTEMP"] = {
        call: secondHighestTemp,
        description: "Returns the second highest from a range of temperatures",
        helpUrl: "https://example.com/help.html",
        result: {
            resultType: Excel.CustomFunctionValueType.number,
            resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
        },
        parameters: [
            {
                name: "temps",
                description: "the temperatures to be compared",
                valueType: Excel.CustomFunctionValueType.number,
                valueDimensionality: Excel.CustomFunctionDimensionality.matrix,
            },
        ],
        options: {
            batch: false,
            stream: false,
        }
    };
    Excel.Script.CustomFunctions["FN"]["TRACKTEMPERATURE"] = {
        call: trackTemperature,
        description: "Streams 25 seconds of temperature history",
        helpUrl: "https://example.com/help.html",
        result: {
            resultType: Excel.CustomFunctionValueType.number,
            resultDimensionality: Excel.CustomFunctionDimensionality.matrix,
        },
        parameters: [
            {
                name: "thermometer ID",
                description: "The thermometer to be measured",
                valueType: Excel.CustomFunctionValueType.string,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
        ],
        options: {
            batch: false,
            stream: true,
        }
    };

    function coPrice(date){
        return Excel.Promise(function(setResult, setError){
            setTimeout(function(){
                setResult((13 + Math.random() * 10).toFixed(2));
            }, 2000);
        });
    }

    function watchOverTime(fields, values, setResult){
        var states = ["AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "DC", "FL", "GA", "ID", "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD", "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ", "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC", "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI", "WY"];
        
        // expect all inputs to be the same-sized columns of single-element arrays, from 1990 to 2015. Sorted by year
        // fields[][0] contains dates. fields[][3] contains the energy source 

        var i = 0;

        function calcNextYear(){
            var year = fields[i][0];
            var output = [[Number(year)]];

            var sums = {};
            for(var j = 0;j < states.length;j++){
                sums[states[j]] = 0;
            }

            while(i < fields.length && fields[i][0] == year){
                if(states.indexOf(fields[i][1] >= 0)){
                    sums[fields[i][1]] += values[i][0];
                }
               
                i++;
            }

            // year has ended
            
            for(var j = 0;j < states.length;j++){
                output.push([Math.round(sums[states[j]])]);
            }
            
            setResult(output);
            
            // reset if needed
            if(year == 2014){
                i = 0;
            }

            setTimeout(calcNextYear, 1000);
        }
        calcNextYear();
        
    }

    function watchOverTime2(fields, values, setResult){
        var states = ["AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "DC", "FL", "GA", "ID", "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD", "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ", "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC", "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI", "WY"];
        
        // expect all inputs to be the same-sized columns of single-element arrays, from 1990 to 2015. Sorted by year
        // fields[][0] contains dates. fields[][3] contains the energy source 

        var i = 0;

        function calcNextYear(){
            var year = fields[i][0];
            var output = [[Number(year)]];

            var sums = {};
            for(var j = 0;j < states.length;j++){
                sums[states[j]] = 0;
            }

            while(i < fields.length && fields[i][0] == year){
                if(states.indexOf(fields[i][1] >= 0)){
                    sums[fields[i][1]] += values[i][0];
                }
               
                i++;
            }

            // year has ended
            
            for(var j = 0;j < states.length;j++){
                output.push([Math.round(sums[states[j]])]);
            }
            
            setResult(output);
            
            // reset if needed
            if(year == 2014){
                i = 0;
            }

            setTimeout(calcNextYear, 1000);
        }
        calcNextYear();
        
    }

    function add42(num1, num2){
        return num1 + num2 + 42;
    }

    Excel.Script.CustomFunctions["CONTOSO"] = {};
    Excel.Script.CustomFunctions["CONTOSO"]["ADD42"] = {
        call: add42,
        description: "Adds two numbers to 42",
        helpUrl: "https://contoso.com/help.html",
        result: {
            resultType: Excel.CustomFunctionValueType.number,
            resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
        },
        parameters: [
            {
                name: "num 1",
                description: "the first number",
                valueType: Excel.CustomFunctionValueType.number,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
            {
                name: "num 2",
                description: "the second number",
                valueType: Excel.CustomFunctionValueType.number,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
        ],
        options: {
            batch: false,
            stream: false,
        }
    };
    Excel.Script.CustomFunctions["FN"] = {};
    Excel.Script.CustomFunctions["FN"]["CO2PRICE"] = {
        call: coPrice,
        description: "Returns the historical price per ton of CO2",
        helpUrl: "https://example.com/help.html",
        result: {
            resultType: Excel.CustomFunctionValueType.number,
            resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
        },
        parameters: [
            {
                name: "date",
                description: "The date to query for the price",
                valueType: Excel.CustomFunctionValueType.number,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
        ],
        options: {
            batch: false,
            stream: false,
        }
    };

    function getPhonePrice(brand, model){
        var price = 1.00;

        switch (model) {
            case 'R9s':
                price = 420.00;
                break;
            case 'A59s':
                price = 430.00;
                break;
            case 'Redmi Note 4X':
                price = 179.00;
                break;
            case 'X9':
                price = 434.00;
                break;
            case 'Redmi 4A':
                price = 115.99;
                break;
            case 'R11':
                price = 528.00;
                break;
            case 'Galaxy J5':
                price = 186.88;
                break;
            case 'Galaxy S8':
                price = 550.00;
                break;
            case 'iPhone 7':
                price = 469.99;
                break;
            case 'Camon CX':
                price = 39.38;
                break;
            case 'Galaxy J1':
                price = 84.99;
                break;
            case 'iPhone SE':
                price = 269.95;
                break;
            case 'Hand':
                price = 320.00;
                break;
            case 'Zoo':
                price = 379.00;
                break;
            case 'Mech':
                price = 315.00;
                break;
            case 'Washer':
                price = 250.00;
                break;
            case 'Coffee':
                price = 299.00;
                break;
            case 'Maker':
                price = 199.00;
                break;
            case 'Wall':
                price = 129.00;
                break;
            case 'E401 Player':
                price = 259.00;
                break;
            case 'E200 Flash':
                price = 319.00;
                break;
            case 'Twoer':
                price = 229.00;
                break;
            default:
                price = 199.00;
        }
          
        return Excel.Promise(function(setResult, setError){
            setTimeout(function(){
                setResult(price);
            }, 250);
        });
    }
    Excel.Script.CustomFunctions["CONTOSO"]["GETPHONEPRICE"] = {
        call: getPhonePrice,
        description: "Queries the web for the latest offered price of the phone model",
        helpUrl: "https://example.com/help.html",
        result: {
            resultType: Excel.CustomFunctionValueType.number,
            resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
        },
        parameters: [
            {
                name: "brand",
                description: "The brand of the phone being queried",
                valueType: Excel.CustomFunctionValueType.string,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
           {
                name: "model",
                description: "The model of the phone being queried",
                valueType: Excel.CustomFunctionValueType.string,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
        ],
        options: {
            batch: false,
            stream: false,
        }
    };

    Excel.Script.CustomFunctions["EPA"] = {};
    Excel.Script.CustomFunctions["EPA"]["DATAOVERTIME"] = {
        call: watchOverTime,
        description: "Visualizes geographical data over time",
        helpUrl: "https://example.com/help.html",
        result: {
            resultType: Excel.CustomFunctionValueType.number,
            resultDimensionality: Excel.CustomFunctionDimensionality.matrix,
        },
        parameters: [
            {
                name: "fields",
                description: "The (increasing, ordered) years and fields for each data point",
                valueType: Excel.CustomFunctionValueType.string,
                valueDimensionality: Excel.CustomFunctionDimensionality.matrix,
            },
           {
                name: "values",
                description: "The values to visualize over time",
                valueType: Excel.CustomFunctionValueType.number,
                valueDimensionality: Excel.CustomFunctionDimensionality.matrix,
            },
        ],
        options: {
            batch: false,
            stream: true,
        }
    };

    Excel.run(function (context) {
        context.workbook.customFunctions.addAll();
        return context.sync().then(function(){

        });
    
    }).catch(function(error){
        console.log("error" + error);
    });
};