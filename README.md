# DocxMerge

Used for merging Microsoft Word documents with priority.


## Installation
This package can be installed from pip using the following:
```
pip install DocxMerge
```

## Usage

```python
from DocxMerge import merge_docs

merge_docs(pivot_doc, final_doc, [file1, file2])
```

Given a document that is a common center point before the branch, a final merged document name, and a list of documents, DocxMerge with merge the documents giving priority in merges to the documents towards the end of the list.
