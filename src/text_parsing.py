from docx import Document
import pickle

# read document
doc_name = 'doc2.docx'

doc = Document(doc_name)

heading_len = 100 # need to determine the optimal

paragraphs = doc.paragraphs
doc_res = dict()
enum = False
i = 0

for i, par in enumerate(paragraphs):
# for par in paragraphs:
    cur_text = par.text
    if not cur_text in ['', '\n', ' ']:
        cur_text = cur_text.rstrip()
        while "  " in cur_text:
            cur_text = cur_text.replace("  ", " ")

        # tables, figures, etc.
        # ...

        # title page (need special flag or something)
        # ...
        
        if enum:
            doc_res[cur_sub]['Enum'] += [cur_text]
            if cur_text[-1] == '.':
                enum = False
        elif cur_text[-1] == ':':
            doc_res[i] = {'Subtitle': cur_text, 'Enum': [paragraphs[i+1].text]}
            enum = True
            cur_sub = i
        elif (cur_text[-1] != '.') and (len(cur_text) < heading_len):
            doc_res[i] = {'Heading': cur_text, 'Text': paragraphs[i+1].text}
            cur_head = i
        else:
            doc_res[cur_head]['Text'] += ' ' + cur_text
            
with open('result.pickle', 'wb') as handle:
    pickle.dump(doc_res, handle, protocol=pickle.HIGHEST_PROTOCOL)
    
print('DONE')