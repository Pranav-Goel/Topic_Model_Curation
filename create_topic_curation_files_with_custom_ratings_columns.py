import configargparse
from pathlib import Path

import numpy as np
import pandas as pd

import xlsxwriter
from wordcloud import WordCloud
import matplotlib.pyplot as plt
from PyPDF2 import PdfFileMerger
from PyPDF2 import PdfFileWriter, PdfFileReader
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

import sys
import os


def rescale_to_probs_renorm(arr):
    return arr/arr.sum(1, keepdims=True)

def create_doc_topic_file_for_annotators(doc_topic,
                                         raw_texts,
                                         outpath,
                                         top_doc_num_per_topic=500):
    '''
    Generates and saves the document-topic file for topic curation.
    
    Input:
    doc_topic: the document-topic probability vector (num_documents X num_topics) [2d numpy array]
    raw_texts: the corresponding raw document texts [list]
    outpath: path to the directory for storing the generated file
    top_doc_num_per_topic: How many of the top documents (using doc_topic) per topic to include in the saved document-topic file. -1 means include all documents. 
    
    '''
    
    out_df = pd.DataFrame()
    num_docs, num_topics = doc_topic.shape
    all_top_doc_inds = set()
    for topic_ind in range(num_topics):
        topic_vals = list(enumerate(list(doc_topic[:, topic_ind])))
        topic_vals = sorted(topic_vals, key=lambda x:x[1])[::-1]
        if top_doc_num_per_topic:
            topic_vals = topic_vals[:top_doc_num_per_topic]
        for ind, _ in topic_vals:
            all_top_doc_inds.add(ind)
    all_top_doc_inds = list(all_top_doc_inds)
    selected_raw_texts = [x for i,x in enumerate(raw_texts) if i in all_top_doc_inds]
    #print(len(all_top_doc_inds))
    out_df['docID'] = [i + 1 for i in range(len(all_top_doc_inds))]
    for topic_ind in range(num_topics):
        out_df['Topic ' + str(topic_ind + 1)] = list(doc_topic[all_top_doc_inds, topic_ind])
    out_df['text'] = selected_raw_texts
    
    out_df.to_excel(Path(outpath) / 'document_topics.xlsx', index=False, float_format='%.3f')
    
def create_topics_file_for_human_labeling_and_rating(topic_word,
                                                     vocab,
                                                     outpath,
                                                     num_top_words = 30,
                                                     custom_cols = {}):
    '''
    Generates and saves the topic-word file for topic curation, can also incorporate creating columns (with drop-down texts where needed) to get ratings/annotations for other things as one creates a label or name for the topic, like rating topic coherence [using custom_cols]
    
    Input:
    topic_word: the topic-word probability vector (num_topics X vocab_size) [2d numpy array]
    vocab: the corresponding words in the vocabulary [list]
    outpath: path to the directory for storing the generated file
    num_top_words: How many of the top words (using topic_word) per topic to include in the saved topic-word file that gets shown. 
    
    custom_cols: a dict where each key is the name of a ratings column to include, and value is a list of values that should be in a drop down for every cell in the column (if the user should choose one of the values to annotate or rate for that column; if the column if meant for free-text input, the list should be empty).
    '''
    
    workbook = xlsxwriter.Workbook(Path(outpath) / 'topic_word.xlsx')
    workbook.formats[0].set_font_size(12)
    worksheet = workbook.add_worksheet()
    worksheet.freeze_panes(1, 0)
    
    # Add a format for the header cells.
    header_format = workbook.add_format({
        'bottom': 5,
        'top':5,
        'font_size':12,
        #'bg_color': '#C6EFCE',
        'bold': True,
        #'text_wrap': True,
        #'valign': 'center',
        'align': 'center',
        #'indent': 1,
    })
    
    
    
    #worksheet.set_row(0, int((topic_word.shape[0]*num_top_words) + topic_word.shape[0] + topic_word.shape[0]))
    
    worksheet.set_column('A:A', 25)
    worksheet.write('A1', 
                    'Topic', 
                    header_format) #just the topic number 'Topic i' followed by the top words of that topic below
    
    worksheet.set_column('B:B', 25)
    worksheet.write('B1', 
                    '', 
                    header_format) #will have excel bars to show the relative strenght of each of the word in that topic for that topic
    
    other_col_start = 67
    #start writing the custom columns to get annotations/ratings on-top of topic curation if needed
    for col in list(custom_cols.keys()):
        worksheet.set_column(chr(other_col_start) + ':' + chr(other_col_start), 40)
        worksheet.write(chr(other_col_start) + '1', 
                        col, 
                        header_format)
        other_col_start += 1
    
    worksheet.set_column(chr(other_col_start) + ':' + chr(other_col_start), 40) 
    worksheet.write(chr(other_col_start) + '1', 
                    'Topic Name',
                    header_format)
    other_col_start += 1
    
    worksheet.set_column(chr(other_col_start) + ':' + chr(other_col_start), 80) 
    worksheet.write(chr(other_col_start) + '1', 
                    'Description',
                    header_format)
    other_col_start += 1
    
    worksheet.set_column(chr(other_col_start) + ':' + chr(other_col_start), 120) 
    worksheet.write(chr(other_col_start) + '1', 
                    'Notes/Comments',
                    header_format)
    

    num_topics, num_words = topic_word.shape
    
    border = workbook.add_format({'top': 2,
                                  'bottom': 2})
    border_plus_highlighting = workbook.add_format({'top': 2,
                                                    'bottom': 2,
                                                    'left': 1,
                                                    'right': 1,
                                                    'bg_color': '#FFFFEC'})
    issue_row = 3
    for k in range(num_topics):
        top_word_inds = np.argsort(list(topic_word[k]))[::-1][:num_top_words]
        top_words = [vocab[i] for i in top_word_inds]
        top_word_probs = [topic_word[k][i] for i in top_word_inds]
        
        worksheet.write('A' + str(issue_row), 'Topic ' + str(k+1), border)
        worksheet.write('B' + str(issue_row), '', border)
        
        #worksheet.set_row(issue_row, None, None, {'collapsed': True})
        
        on_word = 1
        for w, prob in zip(top_words, top_word_probs):
            worksheet.write('A' + str(issue_row + on_word), w)
            worksheet.write('B' + str(issue_row + on_word), prob)
            on_word += 1
        worksheet.conditional_format('B' + str(issue_row + 1) + ':B' + str(issue_row + num_top_words + 1),
                                     {'type': 'data_bar',
                                      'bar_only': True,
                                      'bar_solid': True})
        other_col_start = 67
        for col in list(custom_cols.keys()):
            drop_down_vals = custom_cols[col]
            worksheet.write(chr(other_col_start) + str(issue_row), '', border_plus_highlighting)
            if len(drop_down_vals):
                worksheet.data_validation(chr(other_col_start) + str(issue_row),
                                          {'validate': 'list',
                                           'source': drop_down_vals})
            other_col_start += 1
        
        worksheet.write(chr(other_col_start) + str(issue_row), '', border_plus_highlighting)
        other_col_start += 1
        worksheet.write(chr(other_col_start) + str(issue_row), '', border_plus_highlighting)
        issue_row = issue_row + 2 + num_top_words
    
    workbook.close()

def generate_topic_cloud_image(freq_pairs, outpath):
    wordcloud = WordCloud(max_words=2000,
                          background_color="White",
                          prefer_horizontal=1,
                          relative_scaling=0.75,
                          max_font_size=60,
                          random_state=13)
    wordcloud.generate_from_frequencies(freq_pairs)
    # color_func=lambda *args, **kwargs:(100,100,100) # gray
    color_func=lambda *args, **kwargs:(0,126,157) 
    wordcloud.recolor(color_func=color_func, random_state=3)
    plt.figure()
    plt.imshow(wordcloud, interpolation="bilinear")
    plt.axis("off")
    plt.savefig(outpath,
                bbox_inches='tight')
    plt.clf()
    plt.close()

def write_pdf_with_title(titlestring, pdf_in, pdf_out):
    # https://stackoverflow.com/questions/1180115/add-text-to-existing-pdf-using-python
    # https://stackoverflow.com/questions/9855445/how-to-change-text-font-color-in-reportlab-pdfgen

    packet = io.BytesIO()
    # create a new PDF with Reportlab
    can = canvas.Canvas(packet, pagesize=letter)
    can.setFillColorRGB(0,0,1) #choose your font colour
    can.setFont("Helvetica", 20) #choose your font type and font size
    can.drawString(10, 10, titlestring)
    can.save()
    #move to the beginning of the StringIO buffer
    packet.seek(0)
    new_pdf = PdfFileReader(packet)
    # read your existing PDF
    existing_pdf = PdfFileReader(open(pdf_in, "rb"))
    output = PdfFileWriter()
    # add the "watermark" (which is the new pdf) on the existing page
    page = existing_pdf.getPage(0)
    page.mergePage(new_pdf.getPage(0))
    output.addPage(page)
    # finally, write "output" to a real file
    outputStream = open(pdf_out, "wb")
    output.write(outputStream)
    outputStream.close()
    
def combine_pdfs(pdfs, outfile_name):
    # Solution to merging PDFs into a multi-page PDF document
    # https://stackoverflow.com/questions/3444645/merge-pdf-files
    merger = PdfFileMerger()
    sys.stderr.write("Merging clouds into {}\n".format(outfile_name))
    for pdf in pdfs:
        merger.append(pdf)
    merger.write(outfile_name)
    merger.close()
    
def create_topic_word_clouds_file(topic_word,
                                  vocab,
                                  outdir,
                                  num_words = 50):
    ''' generates the word clouds PDF [combined], control the number of words to use for generating clouds using num_words '''
    pdf_outfile = outdir + "clouds.pdf"
    num_topics, num_words = topic_word.shape
    pdfs = []
    for k in range(num_topics):
        label = 'Topic ' + str(k+1)
        # Temporary files
        image_file_name  = outdir + label.replace(" ", "_") + ".pdf"
        temp_file_name   = outdir + "TEMP_" + label.replace(" ", "_") + ".pdf"
        
        top_word_inds = np.argsort(list(topic_word[k]))[::-1][:num_words]
        top_words = [vocab[i] for i in top_word_inds]
        top_word_probs = [topic_word[k][i] for i in top_word_inds]
        
        word_prob_dic = dict(zip(top_words, top_word_probs))
        
        try:
            # Create image file with cloud
            # sys.stderr.write("Creating cloud image for {} in {}.\n".format(label,temp_file_name))
            generate_topic_cloud_image(word_prob_dic, 
                                       temp_file_name)

            # Create PDF file with topic label, and then remove temporary file
            write_pdf_with_title(label, temp_file_name, image_file_name)
            pdfs.append(image_file_name)
            os.remove(temp_file_name)        

        except Exception as e:
            #sys.stderr.write("---" + str(e) + "---")
            sys.stderr.write("Unable to create cloud image for {} from {} in {}. Skipping.\n".format(label,temp_file_name,image_file_name))
            
    # Combine into one PDF
    combine_pdfs(pdfs, pdf_outfile)

    # Clean up temporary files for the individual cloud images
    for k in range(num_topics):
        label = 'Topic ' + str(k+1)
        image_file_name  = outdir + label.replace(" ", "_") + ".pdf"
        # sys.stderr.write("Removing {}\n".format(image_file_name))
        os.remove(image_file_name)
        


if __name__ == "__main__":
    parser = configargparse.ArgParser()
    parser.add('--topic_word',
               default='example_data/topic_word.npy',
               type=str,
               help='path to .npy file storing the topic-word 2-d numpy array (number of topics X vocab size)')
    parser.add('--doc_topic',
               default='example_data/doc_topic.npy',
               type=str,
               help='path to .npy file storing the document-topic 2-d numpy array (number of documents X number of topics)')
    parser.add('--texts',
               default='example_data/raw_documents.txt',
               type=str,
               help='path to .txt file storing the raw document texts, each line containing one document')
    parser.add('--vocab',
               default='example_data/vocab.txt',
               type=str,
               help='path to .txt file storing the vocabulary, each line containing one term in the vocab')
    
    parser.add('--output',
               default='example_data/outputs/',
               type=str,
               help='path to store the output files generated')
    
    parser.add('--num_top_docs', 
               default=500,
               type=int,
               help='Number of top documents per topic to store in the generated document-topic file used for topic curation; use -1 if all documents are to be contained in the file. For big datasets, specify a number in order to reduce the generated file size.')
    parser.add('--num_top_words', 
               default=30,
               type=int,
               help='Number of top words to show in the generated topic-word excel file (showing bars for relative strength alongside), >=1')
    parser.add('--num_top_words_cloud', 
               default=50,
               type=int,
               help='Number of top words to use when generating topic-word word clouds, >=1')
    
    
    args = parser.parse_args()
    
    # load in the topic-word and document-topic distribution vectors
    topic_word = np.load(args.topic_word)
    doc_topic = np.load(args.doc_topic)
    
    #renormalize to probability vectors (if they are not already)
    if int(topic_word.sum(1).sum()) != topic_word.shape[0]:
        topic_word = rescale_to_probs_renorm(topic_word)
        
    if int(doc_topic.sum(1).sum()) != doc_topic.shape[0]:
        doc_topic = rescale_to_probs_renorm(doc_topic)
        
    # load the raw documents as a list of texts
    raw_texts = open(args.texts).readlines()
    raw_texts = list(map(lambda x:x.rstrip(), raw_texts))
    
    # load the vocabulary as a list of terms
    vocab = open(args.vocab).readlines()
    vocab = list(map(lambda x:x.rstrip(), vocab))
    
    # create output directory if it does not exist
    Path(args.output).mkdir(parents=True, exist_ok=True)

    create_doc_topic_file_for_annotators(doc_topic,
                                         raw_texts,
                                         args.output,
                                         args.num_top_docs)
    print('Document-topic file created.')
    
    '''
    below we have an example of how we could have the topic-word curation file also have some more columns to get
    ratings or annotations, which could have drop down with specified values or open-ended text fields.
    Modify what is passed as the custom_cols arg in the function call below for your own set of custom ratings columns. 
    '''
    custom_columns = {'Coherence': [1, 2, 3], 
                      'Political Polarization': ['IS a polarized issue',
                                                 'MIGHT BE a polarized issue',
                                                 'IS NOT a polarized issue']} #use empty list as value for the key if drop down is not needed
    create_topics_file_for_human_labeling_and_rating(topic_word,
                                                     vocab,
                                                     args.output,
                                                     args.num_top_words,
                                                     custom_cols = custom_columns) 
    print('Topic-word file created.')
    
    print("Creating cloud PDFs...\n")
    create_topic_word_clouds_file(topic_word,
                                  vocab,
                                  args.output,
                                  args.num_top_words_cloud)
    print('\n --- All files created ---')
    
    
    