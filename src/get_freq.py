import sys, argparse
from pathlib import Path
from collections import Counter

import ipdb
import pandas as pd
from box import Box
from ckiptagger import WS

def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument(
            '-c', '--config', dest='config_path',
            default='./config.yaml', type=Path,
            help='the path of config file')
    args = parser.parse_args()
    return vars(args)

def main(config):
    # read data
    df = pd.read_excel(config.data)
    # get all titles, references and add delimeters ('喜見久別的友人 再度帶來物資' => '喜見久別的友人，再度帶來物資')
    titles, references = df['Title'].to_list(), df['Reference'].to_list()
    titles, references = \
        [title.replace(' ', '，') for title in titles], [reference.replace(' ', '，') for reference in references]
    del df

    # word segmentation
    ws = WS(config.tagger_src)
    titles_tokenized, references_tokenized = ws(titles), ws(references)
    del ws
    title_freq, reference_freq = \
        [word for title in titles_tokenized for word in title], [word for reference in references_tokenized for word in reference]
    title_freq, reference_freq = Counter(title_freq), Counter(reference_freq)
    title_freq, reference_freq = \
        title_freq.most_common(len(title_freq)), reference_freq.most_common(len(reference_freq))
    
    # output result
    df_title = pd.DataFrame({
        'Word': [item[0] for item in title_freq], 
        'Frequency': [item[1] for item in title_freq],
    })
    df_reference = pd.DataFrame({
        'Word': [item[0]  for item in reference_freq], 
        'Frequency': [item[1] for item in reference_freq],
    })

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    with pd.ExcelWriter(f'../data/{config.output}') as writer:
        df_title.to_excel(writer, sheet_name='Title Frequency', index=False)
        df_reference.to_excel(writer, sheet_name='Reference Frequency', index=False)

if __name__ == "__main__":
    with ipdb.launch_ipdb_on_exception():
        sys.breakpointhook = ipdb.set_trace
        args = parse_args()
        config = Box.from_yaml(args['config_path'].open())
        main(config)