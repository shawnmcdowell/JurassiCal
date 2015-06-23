(function () {
    var DEFAULT_HEIGHT = 400,
        DEFAULT_WIDTH = 600,
        DEFAULT_BOTTOM_PERCENT = 1 / 3;

    window.grayscale = ['#BEBEBE', '#656565'];
    window.redscale = ['#FF6666', '#B20000'];
    window.bluescale = ['#6666FF', '#0000B2'];
    window.greenscale = ['#66B366', '#006600'];
    window.whitescale = ['#FFFFFF','#FFFFFF'];

    window.colorPalette = "bright";
    window.chartTitle = "Funnel Chart";

    window.outlineThickness = 1; //determines whether there is an outline around each trapezoid. Values 0 (no outline), 1, and 2.8
    window.gapBetweenSlices = 0; //determines whether trapezoids have whitespace in between
    window.curvatureConstant = 3; //adds concavity by increasing gap between y-values of successive trapezoids. Values 0 (no curvatue except for disproportionate data), 1, and 2

    window.labelStyle = "inside";
    window.heightStyle = "variable";
    window.textColor = "#fff";

    window.animationSpeed = 2.5; //Default animationSpeed, updated via settings buttons in Home.js
    window.bottomPct = 1 / 3; //Default bottomPct (ratio of bottom width of funnel to top width), updated via settings buttons in Home.js

    var minData, maxData;

    window.FunnelChart = function (options) {
        /* Parameters:
          data:
            Array containing arrays of categories and engagement in order from greatest expected funnel engagement to lowest.
            I.e. Button loads -> Short link hits
            Ex: [['Button Loads', 1500], ['Button Clicks', 300], ['Subscribers', 150], ['Shortlink Hits', 100]]
          width & height:
            Optional parameters for width & height of chart in pixels, otherwise default width/height are used
          bottomPct:
            Optional parameter that specifies the percent of the total width the bottom of the trapezoid is
            This is used to calculate the slope, so the chart's view can be changed by changing this value
        */

        this.data = options.data;
        minData = this.data[0][1];
        maxData = this.data[0][1];
        this.totalEngagement = 0; //sum of data points in matrix
        for (var i = 0; i < this.data.length; i++) {
            this.totalEngagement += this.data[i][1];
            if (this.data[i][1] < minData) { minData = this.data[i][1]; }
            if (this.data[i][1] > maxData) { maxData = this.data[i][1]; }
        }

        console.log(this.totalEngagement, maxData, minData);

        this.width = typeof options.width !== 'undefined' ? options.width : DEFAULT_WIDTH;
        this.height = typeof options.height !== 'undefined' ? options.height : DEFAULT_HEIGHT;

        this.equalHeight = this.height / this.data.length;
        //this.height = Math.max(this.height, this.data.length * 15)

        var bottomPct = typeof options.bottomPct !== 'undefined' ? options.bottomPct : DEFAULT_BOTTOM_PERCENT;
        this._slope = 2 * this.height / (this.width - bottomPct * this.width);
        console.log("slope " + this._slope);
        this._totalArea = (this.width + bottomPct * this.width) * this.height / 2;
    };

    window.FunnelChart.prototype._getLabel = function (ind) {
        /* Get label of a category at index 'ind' in this.data */
        return this.data[ind][0];
    };

    window.FunnelChart.prototype._getEngagementCount = function (ind) {
        /* Get engagement value of a category at index 'ind' in this.data */
        return this.data[ind][1];
    };

    window.FunnelChart.prototype._createPaths = function () {
        /* Returns an array of points that can be passed into d3.svg.line to create a path for the funnel */
        trapezoids = [];

        function findNextPoints(chart, prevLeftX, prevRightX, prevHeight, dataInd) {
            // reached end of funnel
            if (dataInd >= chart.data.length) return;

            // math to calculate coordinates of the next base
            //Can do things in here with window.curvatureConstant to play around with concavity, but I haven't found a conistent method which works with arbitrary data sets
            area = chart.data[dataInd][1] * chart._totalArea / chart.totalEngagement;
            prevBaseLength = prevRightX - prevLeftX;
            nextBaseLength = Math.sqrt((chart._slope * prevBaseLength * prevBaseLength - 4 * area) / chart._slope);
            nextLeftX = ((prevBaseLength - nextBaseLength) / 2 + prevLeftX);
            nextRightX = (prevRightX - (prevBaseLength - nextBaseLength) / 2);
            
            //Determine height of next trapezoid based on height style
            if (window.heightStyle === "constant") {
                nextHeight = Math.max(prevHeight + chart.equalHeight, prevHeight + 12);
            } else {
                nextHeight = (chart._slope * (prevBaseLength - nextBaseLength) / 2 + prevHeight);
                nextHeight = Math.max(nextHeight, prevHeight + 12) //Give a minimum of 12 height for each trapezoid
                console.log(nextHeight, prevHeight + chart.totalEngagement / 40);
            }

            points = [[nextRightX, nextHeight]];
            points.push([prevRightX, prevHeight]);
            points.push([prevLeftX, prevHeight]);
            points.push([nextLeftX, nextHeight]);
            points.push([nextRightX, nextHeight]);
            trapezoids.push(points);

            findNextPoints(chart, nextLeftX, nextRightX, nextHeight + window.gapBetweenSlices, dataInd + 1);
        }

        findNextPoints(this, 0, this.width, 0, 0);
        return trapezoids;
    };

    window.FunnelChart.prototype.draw = function (elem, speed) {
        var DEFAULT_SPEED = 2.5;
        speed = typeof speed !== 'undefined' ? speed : DEFAULT_SPEED;

        var funnelSvg = d3.select(elem).append('svg:svg')
                  .attr('width', this.width)
                  .attr('height', this.height)
                  .append('svg:g');


        funnelSvg
                    .append('svg:text')
                    .text(window.chartTitle)
                    .attr("x", function (d) { return (this.width); }) //centered text
                    .attr("y", function (d) {
                        return (this.height + 20); //Anchored 4 below mathematical center to fit as well as possible in small trapezoids
                    }) // Average height of bases
                    .attr("text-anchor", "middle")
                    .attr("dominant-baseline", "middle")
                    .attr("fill", "#000");

        // Creates the correct d3 line for the funnel
        var funnelPath = d3.svg.line()
                        .x(function (d) { return d[0]; })
                        .y(function (d) { return d[1]; });

        // Automatically generates colors for each trapezoid in funnel

        console.log(colorPalette);
        if (colorPalette === "bright") {
            var colorScale = d3.scale.category10();
        } else {
            var colorScale = d3.scale.linear()
                .range(colorPalette)
                .domain([minData, maxData]);
        }

        var paths = this._createPaths();

        function drawTrapezoids(funnel, i) {
            var trapezoid = funnelSvg
                            .append('svg:path')
                            .attr('d', function (d) {
                                return funnelPath(
                                    [paths[i][0], paths[i][1], paths[i][2],
                                    paths[i][2], paths[i][1], paths[i][2]]);
                            })
                            .attr('fill', '#fff');

            nextHeight = paths[i][[paths[i].length] - 1];

            var totalLength = trapezoid.node().getTotalLength();

            var transition = trapezoid
                            .transition()
                              .duration(totalLength / speed)
                              .ease("linear")
                              .attr("d", function (d) { return funnelPath(paths[i]); })
                              .attr("fill", function (d) { return colorScale(funnel._getEngagementCount(i)); })
                              .attr('stroke', '#000') //black outline
                              .attr('stroke-width', window.outlineThickness); //with adjustable thickness 0, 1, or 3

            //Place text in trapezoids
            if (window.labelStyle !== "none") {
                var numDigits = (funnel._getEngagementCount(i) + '').replace('.', '').length; //Number of digits in value
                //If label is too long (ARBITRARILY CHOSEN) and value is too long (ARBITRARILY CHOSEN)
                //if (funnel._getLabel(i).length > 30 && numDigits > 10) {
                //    funnelSvg
                //    .append('svg:text')
                //    .text(funnel._getLabel(i).substr(0, 30) + ': ' + funnel._getEngagementCount(i).toPrecision(10))
                //    .attr("x", function (d) { return funnel.width / 2; })
                //    .attr("y", function (d) {
                //        return (paths[i][0][1] + paths[i][1][1]) / 2;
                //    }) // Average height of bases
                //    .attr("text-anchor", "middle")
                //    .attr("dominant-baseline", "middle")
                //    .attr("fill", "#fff");
                //}
                //    //If label is too long
                //else if (funnel._getLabel(i).length > 30) {
                //    funnelSvg
                //    .append('svg:text')
                //    .text(funnel._getLabel(i).substr(0, 30) + ': ' + funnel._getEngagementCount(i))
                //    .attr("x", function (d) { return funnel.width / 2; })
                //    .attr("y", function (d) {
                //        return (paths[i][0][1] + paths[i][1][1]) / 2;
                //    }) // Average height of bases
                //    .attr("text-anchor", "middle")
                //    .attr("dominant-baseline", "middle")
                //    .attr("fill", "#fff");
                //}
                //    //If value is too many digits
                //else if (numDigits > 10) {
                //    funnelSvg
                //    .append('svg:text')
                //    .text(funnel._getLabel(i) + ': ' + funnel._getEngagementCount(i).toPrecision(10))
                //    .attr("x", function (d) { return funnel.width / 2; })
                //    .attr("y", function (d) {
                //        return (paths[i][0][1] + paths[i][1][1]) / 2;
                //    }) // Average height of bases
                //    .attr("text-anchor", "middle")
                //    .attr("dominant-baseline", "middle")
                //    .attr("fill", "#fff");
                //}
                    //No formatting issues
                //else {
                    funnelSvg
                    .append('svg:text')
                    .text(funnel._getLabel(i) + ': ' + funnel._getEngagementCount(i))
                    .attr("x", function (d) { return funnel.width / 2; }) //centered text
                    .attr("y", function (d) {
                        return ((paths[i][0][1] + paths[i][1][1]) / 2) + 4; //Anchored 4 below mathematical center to fit as well as possible in small trapezoids
                    }) // Average height of bases
                    .attr("text-anchor", "middle")
                    .attr("dominant-baseline", "middle")
                    .attr("fill", window.textColor);
                //}
            }

            if (i < paths.length - 1) {
                transition.each('end', function () {
                    drawTrapezoids(funnel, i + 1);
                });
            }
        }

        drawTrapezoids(this, 0);

    };

    window.FunnelChart.prototype.print = function () {
        console.log("I did the thing!");
    }
})();

