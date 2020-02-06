$(document).ready(function () {
    console.log("g1ant script injected on " + (new Date()).format("dd/MM/yyyy hh:mm:ss"));
    function elementClicked(e) {
        var elementIsCrmField = false;
        var mscrmElement = getMscrmSetValueElement(e.target);
        var mscrmFilter = '';
        var mscrmSearch = '';
        if (mscrmElement) {
            elementIsCrmField = mscrmElement.action === 'setvalue';
            mscrmElement = mscrmElement.element;
            var mscrmItem = getMscrmClickedElement(mscrmElement);
            if (mscrmItem) {
                mscrmFilter = mscrmItem.by;
                mscrmSearch = mscrmItem.search;
            }
        }
        else {
            var mscrmClickItem = getMscrmClickedElement(e.target);
            if (mscrmClickItem) {
                mscrmElement = mscrmClickItem.element;
                mscrmFilter = mscrmClickItem.by;
                mscrmSearch = mscrmClickItem.search;
            }
        }

        if (mscrmElement) {
            var element = window.document.createElement('Div');
            element.setAttribute('id', 'MsCrmWizard');
            element.setAttribute('class', 'G1antMsCrmWizardClass');
            element.setAttribute('style', 'display:none');
            element.setAttribute('target_name', mscrmElement.tagName);
            var currentIframe = getIframe(mscrmElement);
            if (currentIframe) {
                var id2 = currentIframe.id ? currentIframe.id : '';
                element.setAttribute('target_iframe_id', id2);
                var title2 = currentIframe.getAttribute('title') ? currentIframe.getAttribute('title') : '';
                element.setAttribute('target_iframe_title', title2);
                var name2 = currentIframe.name ? currentIframe.name : '';
                element.setAttribute('target_iframe_title', name2);
            }
            element.setAttribute('target_by', mscrmFilter);
            element.setAttribute('target_search', mscrmSearch);
            element.setAttribute('data-element-type', elementIsCrmField ? 'setvalue' : 'click');
            window.document.body.appendChild(element);
            console.log((new Date()).format("dd/MM/yyyy hh:mm:ss") + ": Element added. " +
                "Action: " + element.getAttribute('data-element-type') + " " +
                "By: " + element.getAttribute('target_by') + " " +
                "Search: " + element.getAttribute('target_search'));
            inject();
        }
    }

    function getIframe(element) {
        try {
            return element.ownerDocument.defaultView.frameElement;
        }
        catch (e) {
            return null;
        }
    }
    	
	function getMscrmSetValueElement(element) 
	{
        if (element.id.includes('fieldControl-LookupResultsDropdown') ||
            element.tagName.toLowerCase() == 'select' ||
            element.getAttribute('data-id') == 'header_process_decisionmaker.fieldControl-checkbox-toggle' ||
            element.tagName.toLowerCase() == 'textarea' ||
            element.tagName.toLowerCase() == 'input')
		{
			return { element: element, action: 'setvalue' }
		}
		else 
		{
            return { element: element, action: 'click' }
        }
	}
	
    function getMscrmClickedElement(element) {
        while (element.parentElement) {
            if (element.className && element.className === 'ms-crm-InlineLookup-FooterSection-AddAnchor') {
                return { by: 'class', search: element.className, element: element }
            }
            if (element.id) {
                return { by: 'id', search: element.id, element: element }
            }
            if (element.getAttribute('data-id')) {
                return { by: 'data-id', search: element.getAttribute('data-id'), element: element }
            }

            element = element.parentElement;
        }
        return null;
    }

    function inject() {
        var allFrames = [];

        function recursFrames(context) {
            try {
                for (var i = 0; i < context.frames.length; i++) {
                    try { allFrames.push(context.frames[i]); } catch (e) { /*debugger;*/ }
                    try { recursFrames(context.frames[i]); } catch (e) { /*debugger;*/ }
                }
            }
            catch (e) { /*debugger;*/ }
        }
        recursFrames(window);
        allFrames.push(window);
        for (var i = 0; i < allFrames.length; i++) {
            try {
                var currentFrameElements = allFrames[i].document.querySelectorAll(':not([g1antRecorderInjected])');
                for (var j = 0; j < currentFrameElements.length; j++) {
                    try {
                        var currentElement = currentFrameElements[j];
                        if (currentElement.id ||
                            currentElement.getAttribute('data-id') ||
                            (currentElement.className &&
                                (currentElement.className === 'ms-crm-InlineLookup-FooterSection-AddAnchor') // lookup dialog window, 'new' link at the bottom
                            )) {
                            if (currentElement.getAttribute('g1antRecorderInjected') !== 'true') {
                                var tempFunction = null;
                                if (currentElement.tagName.toLowerCase() !== 'button') {

                                    if (currentElement.className === 'navButtonLink') {
                                        try {
                                            currentElement.onclick = function (e) { elementClicked(e); this.ownerDocument.defaultView.DXTools.QuickView.Menu.LoadUrl(this); return false; }
                                        } catch (e) { /*debugger;*/ }
                                    } else

                                        if (currentElement.onclick && currentElement.id) {

                                            tempFunction = currentElement.onclick;
                                            currentElement.onclick = function (e) { elementClicked(e); tempFunction(e); return false; }
                                        }
                                        else {
                                            currentElement.onclick = elementClicked;
                                        }
                                }
                                if (!currentElement.onfocus && currentElement.tagName.toLowerCase() === 'button') {
                                    currentElement.onmousedown = function (e) {
                                        e.target.setAttribute('g1antHoveredMouseDown', 'true');
                                    }
                                    currentElement.onmouseup = function (e) {
                                        e.target.setAttribute('g1antHoveredMouseDown', '');
                                    }
                                    currentElement.onmouseover = function (e) {
                                        e.target.setAttribute('g1antHovered', 'true');
                                    }
                                    currentElement.onmouseleave = function (e) {
                                        e.target.setAttribute('g1antHovered', '');
                                    }
                                    currentElement.onfocus = function (e) {
                                        if (e.target.getAttribute('g1antHovered') === 'true' &&
                                            e.target.getAttribute('g1antHoveredMouseDown') === 'true') {
                                            elementClicked(e);
                                        }
                                    }
                                }
                                currentElement.setAttribute('g1antRecorderInjected', 'true');
                            }
                        }
                    } catch (e) { /*debugger;*/ }
                }
            } catch (e) { /*debugger;*/ }
        }
    }
    inject();
});