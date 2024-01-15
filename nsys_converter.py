from pathlib import Path

import pandas as pd

# 削除対象のColumn
DROP_COLUMNS = ['Avg', 'Med', 'Min', 'Max', 'StdDev']


class NsysConverter:

    def __init__(self, file_path: str, sheet_name: str, save_file_name: str, new_sheet_name: str) -> None:
        self.file_path = Path(file_path)
        self.sheet_name = sheet_name
        self.save_file_name = save_file_name
        self.new_sheet_name = new_sheet_name

    def convert(self, column_name: str) -> None:
        file = pd.read_excel(self.file_path, sheet_name=self.sheet_name, header=0)
        col_edited = []
        columns = file[column_name]
        for data in columns:
            col_edited.append(self._convert_num(data))
        file[column_name] = pd.DataFrame(col_edited)
        # 分かりにくいColumn名を変更
        file = file.rename(columns={'Time': 'Time Rate'})
        # 削除対象のColumnを削除
        file = file.drop(DROP_COLUMNS, axis=1)

        if self.save_file_name is not None:
            # 指定のファイル名でファイルパスを構築
            save_path = self.file_path.parent / f'{self.save_file_name}{self.file_path.suffix}'
        else:
            save_path = f'{self.file_path.stem}_edited{self.file_path.suffix}'
        file.to_excel(save_path, index=False, sheet_name=self.new_sheet_name)

    def _convert_num(self, data: str) -> float:
        value, unit = data.split()
        value = float(value)
        # 単位から数値に変換
        if unit == 's':
            return value
        elif unit == 'ms':
            return value * 0.001
        elif unit == 'μs':
            return value * 0.000001
        else:
            # 上記の単位に当てはまらなければ、そのまま返す
            return value


if __name__ == '__main__':
    from argparse import ArgumentDefaultsHelpFormatter, ArgumentParser
    parser = ArgumentParser(formatter_class=ArgumentDefaultsHelpFormatter)
    parser.add_argument('--file_path', type=str, required=True, help='File path of Nsyight Systems')
    parser.add_argument('--sheet_name', type=str, default='Sheet1', help='Sheet name of result file')
    parser.add_argument('--column_name', type=str, required=True, help='Target column name for convert')
    parser.add_argument('--save_file_name', type=str, default=None, help='File name for edited sheet')
    parser.add_argument('--new_sheet_name', type=str, default='Sheet1', help='Sheet name for edited sheet')
    args = parser.parse_args()
    nc = NsysConverter(args.file_path, args.sheet_name, args.save_file_name, args.new_sheet_name)
    nc.convert(args.column_name)
