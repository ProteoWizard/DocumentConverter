var _ = require("underscore");

var promises = require("./promises");
var Html = require("./html");

exports.imgElement = imgElement;

function imgElement(func) {
    return function(element, messages) {
        return promises.when(func(element)).then(function(result) {
            var attributes = {};
            // disable alt text
            // if (element.altText) {
            //     attributes.alt = element.altText;
            // }
            var imageAttributes = result.attributes ? result.attributes : {};
            _.extend(attributes, imageAttributes);
            var imageElement = [Html.nonFreshElement("img", attributes)];

            // uncomment and return anchor to wrap image element
            // var anchorAttributes =  result.anchorAttributes ? result.anchorAttributes : {};
            // var anchorElement = [Html.freshElement("a", anchorAttributes, imageElement)];
            return imageElement;
        });
    };
}

// Undocumented, but retained for backwards-compatibility with 0.3.x
exports.inline = exports.imgElement;

exports.dataUri = imgElement(function(element) {
    return element.readAsBase64String().then(function(imageBuffer) {
        return {
            src: "data:" + element.contentType + ";base64," + imageBuffer
        };
    });
});
