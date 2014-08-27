olemeta
=======

olemeta is a script to parse OLE files such as MS Office documents (e.g. Word,
Excel), to extract all standard properties present in the OLE file.

It is part of the [python-oletools](http://www.decalage.info/python/oletools) package.

## Usage

	:::text
	olemeta.py <file>

### Example

Checking the malware sample [DIAN_caso-5415.doc](https://malwr.com/analysis/M2I4YWRhM2IwY2QwNDljN2E3ZWFjYTg3ODk4NmZhYmE/):

	:::text
	>olemeta.py DIAN_caso-5415.doc
	
	Properties from SummaryInformation stream:
	- codepage: 1252
	- title: 'Gu\xeda MIPYME para ser emisor electr\xf3nico'
	- subject: ''
	- author: 'OFEyDV'
	- keywords: ''
	- comments: ''
	- template: 'Normal.dotm'
	- last_saved_by: 'clein'
	- revision_number: '13'
	- total_edit_time: 4800L
	- last_printed: datetime.datetime(2006, 6, 7, 14, 4)
	- create_time: datetime.datetime(2009, 3, 30, 14, 18)
	- last_saved_time: datetime.datetime(2014, 5, 14, 12, 45)
	- num_pages: 7
	- num_words: 269
	- num_chars: 1485
	- thumbnail: None
	- creating_application: 'Microsoft Office Word'
	- security: 0
	 
	Properties from DocumentSummaryInformation stream:
	- codepage_doc: 1252
	- category: None
	- presentation_target: None
	- bytes: None
	- lines: 12
	- paragraphs: 3
	- slides: None
	- notes: None
	- hidden_slides: None
	- mm_clips: None
	- scale_crop: False
	- heading_pairs: None
	- titles_of_parts: None
	- manager: None
	- company: 'Servicio de Impuestos Internos'
	- links_dirty: False
	- chars_with_spaces: 1751
	- unused: None
	- shared_doc: False
	- link_base: None
	- hlinks: None
	- hlinks_changed: False
	- version: 786432
	- dig_sig: None
	- content_type: None
	- content_status: None
	- language: None
	- doc_version: None

## How to use olemeta in Python applications	

TODO

--------------------------------------------------------------------------

python-oletools documentation
-----------------------------

- [[Home]]
- [[License]]
- [[Install]]
- [[Contribute]], Suggest Improvements or Report Issues
- Tools:
	- [[olebrowse]]
	- [[oleid]]
	- [[olemeta]]
	- [[oletimes]]
	- [[olevba]]
	- [[pyxswf]]
	- [[rtfobj]] 