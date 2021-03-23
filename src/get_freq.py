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
    with pd.ExcelWriter(f'../data/{config.output}') as writer:
        for sheet in config.sheets:
            # read data
            df = pd.read_excel(config.data, sheet_name=sheet)
            # get all titles, references and add delimeters for title ('喜見久別的友人 再度帶來物資' => '喜見久別的友人，再度帶來物資')
            titles, references = df['Title'].to_list(), df['Reference'].to_list()
            titles, references = \
                [title.strip().replace(' ', '，') for title in titles], [reference.strip() for reference in references]
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
            
            df_out = {
                'Words in Title': [item[0] for item in title_freq], 
                'Frequency of Words in Title': [item[1] for item in title_freq],
                'Words in Reference': [item[0] for item in reference_freq], 
                'Frequency of Words in Reference': [item[1] for item in reference_freq],
            }

            # pad df_out with None otherwise pandas will complain
            min_len = max([len(df_out[key]) for key in df_out])
            df_out = {key: df_out[key]+[None]*(min_len-len(df_out[key])) for key in df_out}
            df_out = pd.DataFrame(df_out)

            # Create a Pandas Excel writer using XlsxWriter as the engine.        
            df_out.to_excel(writer, sheet_name=sheet, index=False)

if __name__ == "__main__":
    with ipdb.launch_ipdb_on_exception():
        sys.breakpointhook = ipdb.set_trace
        args = parse_args()
        config = Box.from_yaml(args['config_path'].open())
        main(config)