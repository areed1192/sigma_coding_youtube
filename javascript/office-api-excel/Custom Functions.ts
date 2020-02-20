/**
 * Calculates the sum of the specified numbers
 * @customfunction
 * @param first First number.
 * @param second Second number.
 * @returns The sum of the numbers.
*/

function addTwo(first: number, second: number): number {
    return first + second;
}

/**
 * Gets the star count for a given Github repository.
 * @customfunction
 * @param {string} userName string name of Github user or organization.
 * @param {string} repoName string name of the Github repository.
 * @return {number} number of stars given to a Github repository.
*/

async function getStarCount(userName: string, repoName:string) {
    try {
        //You can change this URL to any web request you want to work with.
        const url = "https://api.github.com/repos/" + userName + "/" + repoName;
        const response = await fetch(url);

        //Expect that status code is in 200-299 range
        if (!response.ok) {
            throw new Error(response.statusText);
        }

        const jsonResponse = await response.json();
        return jsonResponse.watchers_count;
    } catch (error) {
        return error;
    }
}

/**
 * Returns the current time
 * @returns {string} String with the current time formatted for the current locale.
*/

function currentTime() {
    return new Date().toLocaleTimeString();
}

/**
 * Displays the current time once a second
 * @customfunction
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
*/

function clock(invocation) {
    const timer = setInterval(() => {
        const time = currentTime();
        invocation.setResult(time);
    }, 1000);

    invocation.onCanceled = () => {
        clearInterval(timer);
    };
}

/**
 * Calculates the sum of the specified numbers
 * @customfunction
 * @param first First number.
 * @param second Second number.
 * @param [third] Third number to add. If omitted, third = 0.
 * @returns The sum of the numbers.
*/

function addOpt(first: number, second: number, third?: number): number {
    if (third === null) {
        third = 0;
    }
    return first + second + third;
}

/**
 * Returns the second highest value in a matrixed range of values.
 * @customfunction
 * @param {number[][]} values Multiple ranges of values.
 * @param {number[][]} values2 Multiple ranges of values.
*/

function secondHighest(values, values2) {
    let highest = values[0][0],
        secondHighest = values[0][0];
    for (var i = 0; i < values.length; i++) {
        for (var j = 0; j < values[i].length; j++) {
            if (values[i][j] >= highest) {
                secondHighest = highest;
                highest = values[i][j];
            } else if (values[i][j] >= secondHighest) {
                secondHighest = values[i][j];
            }
        }
    }
    return secondHighest;
}

/**
 * The sum of all of the numbers.
 * @customfunction
 * @param operands A number (such as 1 or 3.1415), a cell address (such as A1 or $E$11), or a range of cell addresses (such as B3:F12)
*/

function AddRepeating(operands: number[][][]): number {
    let total: number = 0;

    operands.forEach((range) => {
        range.forEach((row) => {
            row.forEach((num) => {
                total += num;
            });
        });
    });

    return total;
}

/**
 * Function that gets the address of a cell.
 * @customfunction
 * @param {CustomFunctions.Invocation} invocation Uses the invocation parameter present in each cell.
 * @requiresFont
 * @returns {string} Returns address of cell.
*/

function getAddress(invocation) {
    return invocation.address;
}

/**
 * Get text values that spill down.
 * @customfunction
 * @returns {string[][]} A dynamic array with multiple results.
*/

function spillDown() {
    return [["first"], ["second"], ["third"]];
}

/**
 * Get text values that spill to the right.
 * @customfunction
 * @returns {string[][]} A dynamic array with multiple results.
*/

function spillRight() {
    return [["first", "second", "third"]];
}

/**
 * Get text values that spill both right and down.
 * @customfunction
 * @returns {string[][]} A dynamic array with multiple results.
*/

function spillRectangle() {
    return [["apples", 1, "pounds"], ["oranges", 3, "pounds"], ["pears", 5, "crates"]];
}

/** @CustomFunction
   * @retreives a stock prive for a given symbol
   * @param {string} stock_symbol - The ticker symbol of the stock.
   * @returns A stock price.
*/

async function getStockPrice(stock_symbol: string) {
    try {
        //You can change this URL to any web request you want to work with.
        const url = "https://api.iextrading.com/1.0/tops/last?symbols=" + stock_symbol;
        const response = await fetch(url);

        //Expect that status code is in 200-299 range
        if (!response.ok) {
            throw new Error(response.statusText);
        }

        const jsonResponse = await response.json();
        return jsonResponse[0].price;
    } catch (error) {
        return error;
    }
}
