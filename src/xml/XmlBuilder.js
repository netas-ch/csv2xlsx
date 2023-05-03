/*
    MIT License

    Copyright (c) 2023 Lukas Buchs, netas.ch

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

 */

export class XmlBuilder {
        #xml;
        #rootNode;
        #xmlVersion;
        #xmlStandalone;
        #rootNamespaces;

    constructor(rootName=null, rootNamespace=null, xmlVersion='1.0', xmlStandalone=true) {
        this.#xml = document.implementation.createDocument('', '', null);
        if (rootName) {
            this.#rootNode = this.createAppend(this.#xml, rootName, rootNamespace);
        }

        this.#xmlVersion = xmlVersion;
        this.#xmlStandalone = !!xmlStandalone;
        this.#rootNamespaces = [rootNamespace];
    }

    /**
     * Sets the xml from a string
     * @param {String} xml
     * @returns {undefined}
     */
    setXml(xml) {
        const p = new DOMParser();
        this.#xml = p.parseFromString(xml, 'text/xml');
        this.#rootNode = this.#xml.firstElementChild;
    }

    /**
     * returns the xml as a string
     * @returns {String}
     */
    getXml() {
        const s = new XMLSerializer();
        const str = s.serializeToString(this.#xml);
        return '<?xml version="' + encodeURI(this.#xmlVersion)
                + '" encoding="UTF-8" standalone="' + encodeURI(this.#xmlStandalone ? 'yes': 'no') + '"?>' + "\n" + str;
    }

    /**
     * create and append a element
     * @param {Object|String|null} appendTo
     * @param {String} nodeName
     * @param {String|null} nodeNamespace
     * @param {Object|null} attributes
     * @param {String|null} textContent
     * @returns {Element}
     */
    createAppend(appendTo, nodeName, nodeNamespace, attributes=null, textContent=null) {
        let nde;

        if (appendTo === 'root') {
            appendTo = this.#rootNode;
        }

        // default: same namespace as parent
        if (!nodeNamespace && appendTo && appendTo.namespaceURI) {
            nodeNamespace = appendTo.namespaceURI;
        }

        // add node namespace to root element
        if (this.#rootNode && nodeNamespace !== this.#rootNode.namespaceURI) {
            if (this.#rootNamespaces.indexOf(nodeNamespace) === -1) {

                let namespacePrefix = 'nts' + this.#rootNamespaces.length;

                if (nodeName.split(':').length === 2) {
                    namespacePrefix = nodeName.split(':')[0];
                } else {
                    nodeName = namespacePrefix + ':' + nodeName;
                }

                this.#rootNode.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:" + namespacePrefix, nodeNamespace);
                this.#rootNamespaces.push(nodeNamespace);
            }
        }

        if (!nodeNamespace) {
            nde = this.#xml.createElement(nodeName);
        } else {
            nde = this.#xml.createElementNS(nodeNamespace, nodeName);
        }

        if (appendTo) {
            appendTo.appendChild(nde);
        }

        if (textContent !== null) {
            const txt = this.#xml.createTextNode(textContent);
            nde.appendChild(txt);
        }

        if (attributes !== null) {
            for (const attributeName in attributes) {
                this.setAttribute(nde, attributeName, attributes[attributeName]);
            }
        }

        return nde;
    }

    /**
     * set a attribute for a node
     * @param {Element|String} node or String 'root'
     * @param {String} attributeName
     * @param {String} attributeValue
     * @param {String|null} attributeNamespace
     * @returns {Attr}
     */
    setAttribute(node, attributeName, attributeValue, attributeNamespace=null) {
        let attr;

        if (!attributeNamespace) {
            attr = this.#xml.createAttribute(attributeName);
        } else {

            // add node namespace to root element
            if (this.#rootNode && attributeNamespace !== this.#rootNode.namespaceURI) {
                if (this.#rootNamespaces.indexOf(attributeNamespace) === -1) {

                    let namespacePrefix = 'nts' + this.#rootNamespaces.length;

                    if (attributeName.split(':').length === 2) {
                        namespacePrefix = attributeName.split(':')[0];
                    } else {
                        attributeName = namespacePrefix + ':' + attributeName;
                    }

                    this.#rootNode.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:" + namespacePrefix, attributeNamespace);
                    this.#rootNamespaces.push(attributeNamespace);
                }
            }

            attr = this.#xml.createAttributeNS(attributeNamespace, attributeName);
        }

        if (node === 'root') {
            node = this.#rootNode;
        }

        attr.value = attributeValue;
        node.setAttributeNode(attr);

        return attr;
    }

}