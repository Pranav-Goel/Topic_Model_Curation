Code for creating topic modeling curation files, for annotators to work with topic modeling output in order to label and describe the topics (thus creating human-usable codes, useful for content analysis).

To run the scripts, first set up a conda environment and install the required libraries, run: `conda env create -n <env-name> -f environment.yaml`

The main base script is `create_topic_curation_files.py`, and generated output files that would be used for curation are in `example_data/outputs/`. 

To add custom columns for getting annotations and ratings other than just getting topic labelings, the code in `create_topic_curation_files_with_custom_ratings_columns.py` can be modified and repurposed according to the use case (example output is in `example_data/outputs_with_custom_col_in_topic_word_file`).

NOTE: To run the scripts with the example data in `example_data/`, a `raw_documents.txt` (with one document text per line) must be present (not there because the one that goes with rest of the files is too big).

TO-DO: 

- Further documentation.
- More command line options and flexibility. 
- Custom ratings columns to be provided as an input to the script, likely by having the names and values for those columns in a file. 
