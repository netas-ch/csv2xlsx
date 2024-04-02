/*
        Copyright (c) 2024 Lukas Buchs, netas.ch

    Licensed under the Apache License, Version 2.0 (the "License");
    you may not use this file except in compliance with the License.
    You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

    Unless required by applicable law or agreed to in writing, software
    distributed under the License is distributed on an "AS IS" BASIS,
    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
    See the License for the specific language governing permissions and
    limitations under the License.

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