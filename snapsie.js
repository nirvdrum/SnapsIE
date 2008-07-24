if (!Snapsie)
var Snapsie = new function() {
    // private fields
    var nativeObj;
    
    // private methods
    
    function init() {
        try {
            nativeObj = new ActiveXObject('Snapsie.CoSnapsie');
        }
        catch (e) {
            showException(e);
        }
    }
    
    function showException(e) {
        //throw e;
        alert(e + ', ' + (e.message ? e.message : ""));
    }
    
    function getDrawableElement(inDocument) {
        if (inDocument.compatMode == 'BackCompat') {
            // quirks mode
            return inDocument.getElementsByTagName('body')[0];
        }
        else {
            // standards mode
            return inDocument.documentElement;
        }
    }
    
    /**
     * @param inDocument  the document whose screenshot is being captured
     */
    function getCapturableElement(inDocument) {
        if (inDocument.compatMode == 'BackCompat') {
            // quirks mode
            return inDocument.getElementsByTagName('body')[0];
        }
        else {
            // standards mode
            return inDocument.documentElement;
        }
    }
    
    /**
     * Returns the canonical Windows path for a given path. This means
     * basically replacing any forwards slashes with backslashes.
     *
     * @param path  the path whose canonical form to return
     */
    function getCanonicalPath(path) {
        path = path.replace(/\//g, '\\');
        path = path.replace(/\\\\/g, '\\');
        return path;
    }

    // public methods
    
    /**
     * Saves a screenshot of the current document to a file. If frameId is
     * specified, a screenshot of just the frame is captured instead.
     *
     * @param outputFile  the file to which to save the screenshot
     * @param frameId     the frame to capture; omit to capture entire document
     */
    this.saveSnapshot = function(outputFile, frameId) {
        if (!nativeObj) {
            var e = new Exception('Snapsie was not properly initialized');
            showException(e);
            return;
        }
        
        var drawableElement = getDrawableElement(document);
        var drawableInfo = {
              overflow  : drawableElement.style.overflow
            , scrollLeft: drawableElement.scrollLeft
            , scrollTop : drawableElement.scrollTop
        };
        drawableElement.style.overflow = 'hidden';
        
        var capturableDocument;
        var frameBCR = { left: 0, top: 0 };
        if (arguments.length == 1) {
            capturableDocument = document;
        }
        else {
            var frame = document.getElementById(frameId);
            capturableDocument = frame.document;
            
            // scroll as much of the frame into view as possible
            frameBCR = frame.getBoundingClientRect();
            window.scroll(frameBCR.left, frameBCR.top);
            frameBCR = frame.getBoundingClientRect();
        }
        
        var capturableElement = getCapturableElement(capturableDocument);
        var capturableInfo = {
              overflow  : capturableElement.style.overflow
            , scrollLeft: capturableElement.scrollLeft
            , scrollTop : capturableElement.scrollTop
        };
        capturableElement.style.overflow = 'hidden';
        
        try {
            nativeObj.saveSnapshot(
                getCanonicalPath(outputFile),
                drawableElement.scrollWidth,
                drawableElement.scrollHeight,
                drawableElement.clientWidth,
                drawableElement.clientHeight,
                drawableElement.clientLeft,
                drawableElement.clientTop,
                capturableElement.scrollWidth,
                capturableElement.scrollHeight,
                capturableElement.clientWidth,
                capturableElement.clientHeight,
                frameBCR.left,
                frameBCR.top
            );
        }
        catch (e) {
            showException(e);
        }
        
        // revert, in reverse order in case capturableElement and
        // drawableElement are one and the same
        
        capturableElement.style.overflow = capturableInfo.overflow;
        capturableElement.scrollLeft = capturableInfo.scrollLeft;
        capturableElement.scrollTop = capturableInfo.scrollTop;
        
        drawableElement.style.overflow = drawableInfo.overflow;
        drawableElement.scrollLeft = drawableInfo.scrollLeft;
        drawableElement.scrollTop = drawableInfo.scrollTop;
    }
    
    // initializers
    
    init();
};
