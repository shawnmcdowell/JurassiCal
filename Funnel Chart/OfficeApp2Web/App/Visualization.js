var visualization = (function () {
    "use strict";

    var visualization = {};

    // Sample data:
    visualization.sampleHeaders = [['Stage', 'Percent']];
    visualization.sampleRows = [
        ['Applied', 100],
        ['Interviewed', 70],
        ['Given Offer', 30],
        ['Accepted Offer', 12]];

    

    // Data range validation:
    visualization.rowAndColumnRequirementText = '2 columns and at least 2 rows';

    visualization.isValidRowAndColumnCount = function (rowCount, columnCount) {
        return (rowCount > 1 && columnCount === 2);
    };

    // Creates a visualization, based on passed-in data:
    visualization.createVisualization = function (data) {
        var dat = data.splice(1, data.length); //TODO: determine whether to remove 0th row based on if it's a header or not!!!!!!
        var chart = new FunnelChart({
            data: dat,
            width: 450,
            height: 250,
            bottomPct: window.bottomPct
        });
        chart.draw('#container', window.animationSpeed);
        chart.print();
    };

    return visualization;
})();