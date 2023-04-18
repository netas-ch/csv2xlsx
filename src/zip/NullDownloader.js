/*
 *  https://github.com/Neovici/nullxlsx
 *  Copyright {2020} Neovici
 *
 *  Licensed under the Apache License, Version 2.0 (the "License");
 *  you may not use this file except in compliance with the License.
 *  You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 *  Unless required by applicable law or agreed to in writing, software
 *  distributed under the License is distributed on an "AS IS" BASIS,
 *  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *  See the License for the specific language governing permissions and
 *  limitations under the License.
 */

/**
 * @abstract
 * @unrestricted
 */
export class NullDownloader {
	/**
	 * Creates a new  downloader
	 * @param {string} filename Name of file once generated
	 * @param {string} mimeType The download mime type
	 */

	constructor(filename, mimeType) {
		this._filename = filename;
		this.buffer = null;
		this.lastDownloadBlobUrl = null;
		this._mimeType = mimeType;
	}

	/**
	 * @abstract
	 * @return {ArrayBuffer}  A buffer to download
	 */
	generate() { /* */ }

	/**
	 * Creates an ObjectURL blob containing the generated xlsx
	 * @return {string} ObjectURL to xlsx
	 */
	createDownloadUrl() {
		if (!this.buffer) {
			this.generate();
		}
		const downloadBlob = new Blob([this.buffer], { type: this._mimeType });
		if (this.lastDownloadBlobUrl) {
			window.URL.revokeObjectURL(this.lastDownloadBlobUrl);
		}
		this.lastDownloadBlobUrl = URL.createObjectURL(downloadBlob);
		return this.lastDownloadBlobUrl;
	}

	/**
	 * Create download link object (or update existing)
	 * @param {(string|HTMLAnchorElement)} linkText Existing link object or text to set on new link
	 * @return {!Element} Link object
	 */
	createDownloadLink(linkText) {
		const link = linkText instanceof HTMLAnchorElement ? linkText : document.createElement('a');
		if (typeof linkText === 'string') {
			link.innerHTML = linkText;
		}
		link.href = this.createDownloadUrl();
		link.download = this._filename;
                link.innerText = this._filename;
		
		return link;
	}
}
