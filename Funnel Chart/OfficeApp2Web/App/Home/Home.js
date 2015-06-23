/// <reference path="../App.js" />
/// <reference path="../Visualization.js" />

(function () {
    "use strict";

    var binding = null;

    // Default displayed data
    //TODO: WHY THE HECK IS THIS CHANGING TO JUST DEFAULTDATA[0] AFTER THE FIRST PASS (CAUSING IT TO BE INVALID DATA AND SETTINGS DON'T WORK)??
    var defaultData = [['Category', 'Number'], ['Clicks', 768], ['Free Downloads', 455], ['Purchases', 211], ['Repeat Purchases', 134]];

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            displayExistingData();

            var level0 = document.getElementsByClassName('level0');
            var level1 = document.getElementsByClassName('level1');
            var level2 = document.getElementsByClassName('level2');
            var level3 = document.getElementsByClassName('level3');
            var level4 = document.getElementsByClassName('level4');

            //When clicking anywhere that isn't a submenu option or the passive parts of a flyout (unclickable body or header), close flyouts
            $(document).on('click', function (event) {
                var target = $(event.target);
                if (!target.hasClass('headerText') && !target.hasClass('body') && !target.hasClass('submenuoption')) {
                    var flyouts = document.getElementsByClassName('flyout');
                    for (var i = 0; i < flyouts.length; i++) {
                        flyouts[i].parentNode.removeChild(flyouts[i]);
                    }
                }
            })

            //Click menu button to display dropdown submenu
            document.getElementById('menu').onclick = function () {
                document.getElementById('submenu').style.display = 'block';
                //Restore level1 display and all other levels hidden (in case this was changed elsewhere)
                for (var i = 0; i < level1.length; i++) {
                    level1[i].style.display = "block";
                }
                for (var i = 0; i < level0.length; i++) {
                    level0[i].style.display = "none";
                }
                for (var i = 0; i < level2.length; i++) {
                    level2[i].style.display = "none";
                }
                for (var i = 0; i < level3.length; i++) {
                    level3[i].style.display = "none";
                }
                for (var i = 0; i < level4.length; i++) {
                    level4[i].style.display = "none";
                }
            };

            //Click datamenu button to display dropdown submenu
            document.getElementById('datamenu').onclick = function () {
                document.getElementById('submenu').style.display = 'block';
                //Restore level1 display and all other levels hidden (in case this was changed elsewhere)
                for (var i = 0; i < level0.length; i++) {
                    level0[i].style.display = "block";
                }
                for (var i = 0; i < level1.length; i++) {
                    level1[i].style.display = "none";
                }
                for (var i = 0; i < level2.length; i++) {
                    level2[i].style.display = "none";
                }
                for (var i = 0; i < level3.length; i++) {
                    level3[i].style.display = "none";
                }
                for (var i = 0; i < level4.length; i++) {
                    level4[i].style.display = "none";
                }
            };

            //Click a submenu level1 option to show next level
            document.getElementById('colorButton').onclick = function () {
                for(var i=0; i<level1.length; i++){
                    level1[i].style.display = "none";
                }
                for (var i = 0; i < level2.length; i++) {
                    level2[i].style.display = "block";
                }
            };
            document.getElementById('shapeButton').onclick = function () {
                for (var i = 0; i < level1.length; i++) {
                    level1[i].style.display = "none";
                }
                for (var i = 0; i < level3.length; i++) {
                    level3[i].style.display = "block";
                }
            };
            document.getElementById('miscButton').onclick = function () {
                for (var i = 0; i < level1.length; i++) {
                    level1[i].style.display = "none";
                }
                for (var i = 0; i < level4.length; i++) {
                    level4[i].style.display = "block";
                }
            };

            //Click a submenu options to display the proper flyout
            document.getElementById('sampleButton').onclick = function () { insertSampleData(); };
            document.getElementById('bindButton').onclick = function () { bindToExistingData(); };

            document.getElementById('titleButton').onclick = function () { showTitleFlyout(); };
            document.getElementById('chartColorButton').onclick = function () { showColorFlyout(); };
            document.getElementById('labelsButton').onclick = function () { showLabelFlyout(); };
            document.getElementById('animationButton').onclick = function () {
                showAnimationFlyout();
                for (var i = 0; i < level1.length; i++) {
                    level1[i].style.display = "none";
                }
            };
            document.getElementById('slopeButton').onclick = function () { showSlopeFlyout(); };
            document.getElementById('viewButton').onclick = function () { showViewFlyout(); };
            document.getElementById('outlineButton').onclick = function () { showOutlineFlyout(); };
            document.getElementById('gapButton').onclick = function () { showGapFlyout(); };


            //When clicking off of the menu button itself, close dropdown submenu
            document.getElementById('content-main').onclick = function (e) {
                var textbox = $(e.target).closest('.textBox');
                if (e.target != document.getElementById('menu') &&
                    e.target != document.getElementById('datamenu') &&
                    e.target != document.getElementById('textBox') &&
                    e.target.tagName != "INPUT" &&
                    !$(e.target).hasClass('level1')) {
                        document.getElementById('submenu').style.display = 'none';
                }
            };
        });
    };

    //TODO: MUST DELETE ANY BINDING WHICH ALREADY EXISTS!!!!
    function insertSampleData() {
        var sampleData = new Office.TableData(
            visualization.sampleRows,
            visualization.sampleHeaders);
        Office.context.document.setSelectedDataAsync(sampleData,
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    app.showNotification('Could not insert sample data',
                        'Please choose a different selection range.');
                } else {
                    Office.context.document.bindings.addFromSelectionAsync(
                        Office.BindingType.Table, { id: app.bindingID },
                        function (asyncResult) {
                            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                                app.showNotification('Error binding data');
                            } else {
                                window.location.href = '../Home/Home.html';
                            }
                        }
                    );
                }
            }
        );
    }

    function bindToExistingData() {
        Office.context.document.bindings.addFromSelectionAsync(
            Office.BindingType.Matrix, { id: app.bindingID }, function (result) {
                var isValid = (result.status == Office.AsyncResultStatus.Succeeded) && visualization.isValidRowAndColumnCount(result.value.rowCount, result.value.columnCount);
                if (isValid) {
                    window.location.href = '../Home/Home.html';
                } else {
                    app.showNotification('Invalid data selected',
                        'Please make a different selection, and ensure that you selected a table or range with ' + visualization.rowAndColumnRequirementText);
                }
            }
        );
    }


    function showHideDiv(div) {
        if (div.style.display == 'none') {
            div.style.display = 'block';
        } else {
            div.style.display = 'none';
        }
    }

    function createHeader(text) {
        var header = document.createElement("div");
        header.className = "header";
        var headerText = document.createElement("div");
        headerText.className = "headerText";
        headerText.innerText = text;
        header.appendChild(headerText);
        return header;
    }

    function showTitleFlyout() {
        var flyout = document.createElement("div");
        flyout.className = "flyout";

        var header = createHeader("Enter a Custom Title");

        var body = document.createElement("div");
        body.className = "body";

        var textBox = document.createElement("input");
        textBox.id = "textBox";
        textBox.innerText = window.chartTitle;

        //Append each color option to the body
        body.appendChild(textBox);

        //Append header and body to the flyout
        flyout.appendChild(header);
        flyout.appendChild(body);

        //Append the flyout to the page so it is displayed
        document.getElementById("flyoutContainer").appendChild(flyout);

        $("#textBox").keyup(function (e) {
            if (e.keyCode === 13) {
                if (window.chartTitle !== this.value) {
                    window.chartTitle = this.value;
                    displayDataForBinding(binding);
                }
                console.log(this.value);

            }
        })

        //document.getElementById('textBox').onkeyup = function () {
        //    window.chartTitle = textBox.innerHTML;
        //}
        //document.getElementById('textBox').onkeypress = function () {
        //    window.chartTitle = textBox.innerHTML;
        //}
    }

    function showColorFlyout() {
        var flyout = document.createElement("div");
        flyout.className = "flyout";

        var header = createHeader("Select a Color Palette");

        var body = document.createElement("div");
        body.className = "body";

        //Option for bright colors
        var bright = document.createElement("div");
        bright.className = "settingOption";
        var colorSwatch = document.createElement("div");
        colorSwatch.className = "colorSwatch";
        //Create 3 colors to fit inside colorSwatch box
        var color1 = document.createElement("div");
        color1.className = "subSwatch red";
        colorSwatch.appendChild(color1);
        var color2 = document.createElement("div");
        color2.className = "subSwatch blue";
        colorSwatch.appendChild(color2);
        var color3 = document.createElement("div");
        color3.className = "subSwatch orange";
        colorSwatch.appendChild(color3);
        //Text label
        var text = document.createElement("div");
        text.className = "text";
        text.innerText = "Bright Colors";
        bright.appendChild(colorSwatch);
        bright.appendChild(text);
        console.log(window.colorPalette);
        //When bright colors is clicked, reload graph with category10 as the color palette
        bright.addEventListener("click", function () {
            if (window.colorPalette !== "bright") {
                window.colorPalette = "bright";
                window.textColor = "#fff";
                displayDataForBinding(binding);
            }
        });

        //Option for grayscale
        var gray = document.createElement("div");
        gray.className = "settingOption";
        var colorSwatch = document.createElement("div");
        colorSwatch.className = "colorSwatch gray";
        var text = document.createElement("div");
        text.className = "text";
        text.innerText = "Gray Scale";
        gray.appendChild(colorSwatch);
        gray.appendChild(text);
        //When grayscale is clicked, reload graph with grayscale as the color palette
        gray.addEventListener("click", function () {
            if (window.colorPalette !== window.grayscale) {
                window.colorPalette = window.grayscale;
                window.textColor = "#fff";
                displayDataForBinding(binding);
            }
        });

        //Option for redscale
        var red = document.createElement("div");
        red.className = "settingOption";
        var colorSwatch = document.createElement("div");
        colorSwatch.className = "colorSwatch red";
        var text = document.createElement("div");
        text.className = "text";
        text.innerText = "Red Scale";
        red.appendChild(colorSwatch);
        red.appendChild(text);
        //When redscale is clicked, reload graph with redscale as the color palette
        red.addEventListener("click", function () {
            if (window.colorPalette !== window.redscale) {
                window.colorPalette = window.redscale;
                window.textColor = "#fff";
                displayDataForBinding(binding);
            }
        });

        //Option for greenscale
        var green = document.createElement("div");
        green.className = "settingOption";
        var colorSwatch = document.createElement("div");
        colorSwatch.className = "colorSwatch green";
        var text = document.createElement("div");
        text.className = "text";
        text.innerText = "Green Scale";
        green.appendChild(colorSwatch);
        green.appendChild(text);
        //When grayscale is clicked, reload graph with grayscale as the color palette
        green.addEventListener("click", function () {
            if (window.colorPalette !== window.greenscale) {
                window.colorPalette = window.greenscale;
                window.textColor = "#fff";
                displayDataForBinding(binding);
            }
        });

        //Option for bluescale
        var blue = document.createElement("div");
        blue.className = "settingOption";
        var colorSwatch = document.createElement("div");
        colorSwatch.className = "colorSwatch blue";
        var text = document.createElement("div");
        text.className = "text";
        text.innerText = "Blue Scale";
        blue.appendChild(colorSwatch);
        blue.appendChild(text);
        //When blue is clicked, reload graph with bluescale as the color palette
        blue.addEventListener("click", function () {
            if (window.colorPalette !== window.bluescale) {
                window.colorPalette = window.bluescale;
                window.textColor = "#fff";
                displayDataForBinding(binding);
            }
        });

        //Option for whiteout
        var white = document.createElement("div");
        white.className = "settingOption";
        var colorSwatch = document.createElement("div");
        colorSwatch.className = "colorSwatch white";
        var text = document.createElement("div");
        text.className = "text";
        text.innerText = "No Color";
        white.appendChild(colorSwatch);
        white.appendChild(text);
        //When white is clicked, reload graph with whitescale as the color palette
        white.addEventListener("click", function () {
            if (window.colorPalette !== window.whitescale) {
                window.colorPalette = window.whitescale;
                //Must have an outline for white only; default to 1 only if there is no current outline
                if (window.outlineThickness === 0) {
                    window.outlineThickness = 1;
                }
                window.textColor = "#000";
                displayDataForBinding(binding);
            }
        });

        //Append each color option to the body
        body.appendChild(bright);
        body.appendChild(gray);
        body.appendChild(red);
        body.appendChild(green);
        body.appendChild(blue);
        body.appendChild(white);

        //Append header and body to the flyout
        flyout.appendChild(header);
        flyout.appendChild(body);

        //Append the flyout to the page so it is displayed
        document.getElementById("flyoutContainer").appendChild(flyout);
    }

    function showLabelFlyout() {
        var flyout = document.createElement("div");
        flyout.className = "flyout";

        var header = createHeader("Select a Label Style");
        flyout.appendChild(header);

        var body = document.createElement("div");
        body.className = "body";

        //Option for inside
        var inside = document.createElement("div");
        inside.className = "settingOption";
        var text = document.createElement("div");
        text.className = "text";
        text.innerText = "Inside Chart";
        inside.appendChild(text);
        //When inside is clicked, reload graph with inside labels
        inside.addEventListener("click", function () {
            if (window.labelStyle !== "inside") {
                window.labelStyle = "inside";
                displayDataForBinding(binding);
            }
        });

        //Option for outside
        var outside = document.createElement("div");
        outside.className = "settingOption";
        var text = document.createElement("div");
        text.className = "text";
        text.innerText = "Outside Chart";
        outside.appendChild(text);
        //When outside is clicked, reload graph with outside labels
        outside.addEventListener("click", function () {
            if (window.labelStyle !== "outside") {
                window.labelStyle = "outside";
                displayDataForBinding(binding);
            }
        });

        //Option for color-coordinated key
        var key = document.createElement("div");
        key.className = "settingOption";
        var text = document.createElement("div");
        text.className = "text";
        text.innerText = "Colored Key";
        key.appendChild(text);
        //When outside is clicked, reload graph with outside labels
        key.addEventListener("click", function () {
            if (window.labelStyle !== "key") {
                window.labelStyle = "key";
                displayDataForBinding(binding);
            }
        });

        //Option for no labels
        var noLabels = document.createElement("div");
        noLabels.className = "settingOption";
        var text = document.createElement("div");
        text.className = "text";
        text.innerText = "None";
        noLabels.appendChild(text);
        //When outside is clicked, reload graph with outside labels
        noLabels.addEventListener("click", function () {
            if (window.labelStyle !== "none") {
                window.labelStyle = "none";
                displayDataForBinding(binding);
            }
        });

        //Append each label option to body
        body.appendChild(inside);
        body.appendChild(outside);
        body.appendChild(key);
        body.appendChild(noLabels);

        //Append header and body to the flyout
        flyout.appendChild(header);
        flyout.appendChild(body);

        document.addEventListener("click", function () {

        });

        //Append the flyout to the page so it is displayed
        document.getElementById("flyoutContainer").appendChild(flyout);

    }

    function showAnimationFlyout() {
        var flyout = document.createElement("div");
        flyout.className = "flyout";

        var header = createHeader("Select an Animation Speed");
        header.className = "header";

        var body = document.createElement("div");
        body.className = "body";

        //Option for no animation (really just incredibly fast)
        var noAnimation = document.createElement("div");
        noAnimation.className = "settingOption";
        var text = document.createElement("div");
        text.className = "text";
        text.innerText = "None";
        noAnimation.appendChild(text);
        //When slow is clicked, reload graph with slow animation
        noAnimation.addEventListener("click", function () {
            if (window.animationSpeed !== 1000) {
                window.animationSpeed = 1000;
                displayDataForBinding(binding);
            }
        });

        //Option for slow animation
        var slow = document.createElement("div");
        slow.className = "settingOption";
        var text = document.createElement("div");
        text.className = "text";
        text.innerText = "Slow";
        slow.appendChild(text);
        //When slow is clicked, reload graph with slow animation
        slow.addEventListener("click", function () {
            if (window.animationSpeed !== 1) {
                window.animationSpeed = 1;
                displayDataForBinding(binding);
            }
        });

        //Option for medium animation
        var medium = document.createElement("div");
        medium.className = "settingOption";
        var text = document.createElement("div");
        text.className = "text";
        text.innerText = "Moderate";
        medium.appendChild(text);
        //When slow is clicked, reload graph with slow animation
        medium.addEventListener("click", function () {
            if (window.animationSpeed !== 2.5) {
                window.animationSpeed = 2.5;
                displayDataForBinding(binding);
            }
        });

        //Option for fast animation
        var fast = document.createElement("div");
        fast.className = "settingOption";
        var text = document.createElement("div");
        text.className = "text";
        text.innerText = "Fast";
        fast.appendChild(text);
        //When slow is clicked, reload graph with slow animation
        fast.addEventListener("click", function () {
            if (window.animationSpeed !== 6) {
                window.animationSpeed = 6;
                displayDataForBinding(binding);
            }
        });

        //Append each animation option to body
        body.appendChild(noAnimation);
        body.appendChild(slow);
        body.appendChild(medium);
        body.appendChild(fast);

        //Append header and body to the flyout
        flyout.appendChild(header);
        flyout.appendChild(body);

        //Append the flyout to the page so it is displayed
        document.getElementById("flyoutContainer").appendChild(flyout);

    }

    function showSlopeFlyout() {
        var flyout = document.createElement("div");
        flyout.className = "flyout";

        var header = createHeader("Select Base Width");
        flyout.appendChild(header);

        var body = document.createElement("div");
        body.className = "body";

        //Option for wide base (close to rectangular)
        var veryWide = document.createElement("div");
        veryWide.className = "settingOption";
        var colorSwatch = document.createElement("div");
        colorSwatch.className = "colorSwatch blue";
        var text = document.createElement("div");
        text.className = "text";
        text.innerText = "Very Wide";
        veryWide.appendChild(colorSwatch);
        veryWide.appendChild(text);
        //When wide is clicked, reload graph with wide base
        veryWide.addEventListener("click", function () {
            if (window.bottomPct !== 999 / 1000) {
                window.bottomPct = 999 / 1000;
                displayDataForBinding(binding);
            }
        });

        //Option for mid-wide base (high slope sides)
        var midWide = document.createElement("div");
        midWide.className = "settingOption";
        var colorSwatch = document.createElement("div");
        colorSwatch.className = "colorSwatch blue";
        var text = document.createElement("div");
        text.className = "text";
        text.innerText = "Wide";
        midWide.appendChild(colorSwatch);
        midWide.appendChild(text);
        //When wide is clicked, reload graph with mid-wide base
        midWide.addEventListener("click", function () {
            if (window.bottomPct !== 1 / 3) {
                window.bottomPct = 1 / 3;
                displayDataForBinding(binding);
            }
        });

        //Option for mid-narrow base (low slope sides)
        var midNarrow = document.createElement("div");
        midNarrow.className = "settingOption";
        var colorSwatch = document.createElement("div");
        colorSwatch.className = "colorSwatch blue";
        var text = document.createElement("div");
        text.className = "text";
        text.innerText = "Narrow";
        midNarrow.appendChild(colorSwatch);
        midNarrow.appendChild(text);
        //When wide is clicked, reload graph with mid-wide base
        midNarrow.addEventListener("click", function () {
            if (window.bottomPct !== 1 / 6) {
                window.bottomPct = 1 / 6;
                displayDataForBinding(binding);
            }
        });

        //Option for narrow base (very high slope sidescomes to point at bottom)
        var zeroWidth = document.createElement("div");
        zeroWidth.className = "settingOption";
        var colorSwatch = document.createElement("div");
        colorSwatch.className = "colorSwatch blue";
        var text = document.createElement("div");
        text.className = "text";
        text.innerText = "Very Narrow";
        zeroWidth.appendChild(colorSwatch);
        zeroWidth.appendChild(text);
        //When wide is clicked, reload graph with mid-wide base
        zeroWidth.addEventListener("click", function () {
            if (window.bottomPct !== 1 / 1000) {
                window.bottomPct = 1 / 1000;
                displayDataForBinding(binding);
            }
        });

        //Append each slope option to body
        body.appendChild(veryWide);
        body.appendChild(midWide);
        body.appendChild(midNarrow);
        body.appendChild(zeroWidth);

        //Append header and body to the flyout
        flyout.appendChild(header);
        flyout.appendChild(body);

        //Append the flyout to the page so it is displayed
        document.getElementById("flyoutContainer").appendChild(flyout);

    }

    function showViewFlyout() {
        var flyout = document.createElement("div");
        flyout.className = "flyout";

        var header = createHeader("Select a View");
        header.className = "header";

        var body = document.createElement("div");
        body.className = "body";

        //Option for variable height (default)
        var defaultView = document.createElement("div");
        defaultView.className = "settingOption";
        var text = document.createElement("div");
        text.className = "text";
        text.innerText = "Variable Heights";
        defaultView.appendChild(text);
        //When slow is clicked, reload graph with slow animation
        defaultView.addEventListener("click", function () {
            if (window.heightStyle !== "variable") {
                window.heightStyle = "variable";
                displayDataForBinding(binding);
            }
        });

        //Option for constant heights
        var constantHeight = document.createElement("div");
        constantHeight.className = "settingOption";
        var text = document.createElement("div");
        text.className = "text";
        text.innerText = "Constant Heights";
        constantHeight.appendChild(text);
        //When slow is clicked, reload graph with slow animation
        constantHeight.addEventListener("click", function () {
            if (window.heightStyle !== "constant") {
                window.heightStyle = "constant";
                displayDataForBinding(binding);
            }
        });

        //Append each view option to body
        body.appendChild(defaultView);
        body.appendChild(constantHeight);

        //Append header and body to the flyout
        flyout.appendChild(header);
        flyout.appendChild(body);

        document.getElementById("flyoutContainer").appendChild(flyout);

    }

    function showOutlineFlyout() {
        var flyout = document.createElement("div");
        flyout.className = "flyout";

        var header = createHeader("Select an Outline Style");
        flyout.appendChild(header);

        var body = document.createElement("div");
        body.className = "body";

        //Option for no gap
        var noOutline = document.createElement("div");
        noOutline.className = "settingOption";
        var text = document.createElement("div");
        text.className = "text";
        text.innerText = "No Outline";
        noOutline.appendChild(text);
        //When slow is clicked, reload graph with slow animation
        noOutline.addEventListener("click", function () {
            if (window.outlineThickness !== 0) {
                window.outlineThickness = 0;
                displayDataForBinding(binding);
            }
        });

        //Option for fast animation
        var thinOutline = document.createElement("div");
        thinOutline.className = "settingOption";
        var text = document.createElement("div");
        text.className = "text";
        text.innerText = "Thin Outline";
        thinOutline.appendChild(text);
        //When slow is clicked, reload graph with slow animation
        thinOutline.addEventListener("click", function () {
            if (window.outlineThickness !== 1) {
                window.outlineThickness = 1;
                displayDataForBinding(binding);
            }
        });

        //Option for fast animation
        var thickOutline = document.createElement("div");
        thickOutline.className = "settingOption";
        var text = document.createElement("div");
        text.className = "text";
        text.innerText = "Thick Outline";
        thickOutline.appendChild(text);
        //When slow is clicked, reload graph with slow animation
        thickOutline.addEventListener("click", function () {
            if (window.outlineThickness !== 2.8) {
                window.outlineThickness = 2.8;
                displayDataForBinding(binding);
            }
        });

        //Append each view option to body
        body.appendChild(noOutline);
        body.appendChild(thinOutline);
        body.appendChild(thickOutline);

        //Append header and body to the flyout
        flyout.appendChild(header);
        flyout.appendChild(body);

        document.getElementById("flyoutContainer").appendChild(flyout);

    }

    function showGapFlyout() {
        var flyout = document.createElement("div");
        flyout.className = "flyout";

        var header = createHeader("Select Spacing");
        flyout.appendChild(header);

        var body = document.createElement("div");
        body.className = "body";

        //Option for no gap
        var noGap = document.createElement("div");
        noGap.className = "settingOption";
        var text = document.createElement("div");
        text.className = "text";
        text.innerText = "No Gap";
        noGap.appendChild(text);
        //When slow is clicked, reload graph with slow animation
        noGap.addEventListener("click", function () {
            if (window.gapBetweenSlices !== 0) {
                window.gapBetweenSlices = 0;
                displayDataForBinding(binding);
            }
        });

        //Option for fast animation
        var smallGap = document.createElement("div");
        smallGap.className = "settingOption";
        var text = document.createElement("div");
        text.className = "text";
        text.innerText = "Narrow Gap";
        smallGap.appendChild(text);
        //When slow is clicked, reload graph with slow animation
        smallGap.addEventListener("click", function () {
            if (window.gapBetweenSlices !== 4) {
                window.gapBetweenSlices = 4;
                displayDataForBinding(binding);
            }
        });

        //Option for fast animation
        var largeGap = document.createElement("div");
        largeGap.className = "settingOption";
        var text = document.createElement("div");
        text.className = "text";
        text.innerText = "Wide Gap";
        largeGap.appendChild(text);
        //When slow is clicked, reload graph with slow animation
        largeGap.addEventListener("click", function () {
            if (window.gapBetweenSlices !== 10) {
                window.gapBetweenSlices = 10;
                displayDataForBinding(binding);
            }
        });

        //Append each view option to body
        body.appendChild(noGap);
        body.appendChild(smallGap);
        body.appendChild(largeGap);

        //Append header and body to the flyout
        flyout.appendChild(header);
        flyout.appendChild(body);

        document.getElementById("flyoutContainer").appendChild(flyout);

    }


    //Checks if a binding exists, and displays the corresponding funnel chart if so
    function displayExistingData() {
        Office.context.document.bindings.getByIdAsync(
            app.bindingID,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    binding = result.value;
                    displayDataForBinding(binding);
                    // And bind a change-event handler to the binding:
                    binding.addHandlerAsync(
                        Office.EventType.BindingDataChanged,
                        function () {
                            displayDataForBinding(binding);
                        }
                    );
                } else {
                    displayDataHelper(defaultData);
                }
            });
    }

    //Reads in data from selected matrix
    function displayDataForBinding(binding) {
        //TODO: ADD CHECK FOR VALIDITY SO THAT YOU DON'T REFRESH IN A WAY THAT WILL MESS UP THE TABLE!
        if (binding) {
            binding.getDataAsync({ coercionType: Office.CoercionType.Matrix, valueFormat: Office.ValueFormat.Unformatted, filterType: Office.FilterType.OnlyVisible },
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        displayDataHelper(result.value);
                    } else {
                        $('#data-display').html(
                            '<div class="notice">' +
                            '    <h2>Error fetching data!</h2>' +
                            '    <a href="../DataBinding/DataBinding.html">' +
                            '        <b>Bind to a different range?</b>' +
                            '    </a>' +
                            '</div>');
                    }
                }
            );
        } else { //No binding exists
            //$('#data-display').html(
            //                '<div class="notice">' +
            //                '    <h2>Error fetching data!</h2>' +
            //                '    <a href="../DataBinding/DataBinding.html">' +
            //                '        <b>Bind to a different range?</b>' +
            //                '    </a>' +
            //                '</div>');
            //app.showNotification('Could not find valid data binding',
            //            'Please select data and choose "Bind to Existing Data".');
            //return;
            displayDataHelper(defaultData);
        }
    }

    //Takes in selected matrix data, checks valid row/col counts, calls visualization.createVisualization, appends result to #data-display
    function displayDataHelper(data) {
        var rowCount = data.length;
        var columnCount = (data.length > 0) ? data[0].length : 0;
        if (!visualization.isValidRowAndColumnCount(rowCount, columnCount)) {
            $('#data-display').html(
                '<div class="notice">' +
                '    <h2>Not enough data!</h2>' +
                '    <p>The range must contain ' + visualization.rowAndColumnRequirementText + '.</p>' +
                '    <a href="../DataBinding/DataBinding.html">' +
                '        <b>Choose a different range?</b>' +
                '    </a>' +
                '</div>');
            return;
        }

        $('#container').empty();
        visualization.createVisualization(data);
    }
})();