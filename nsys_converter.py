import pandas as pd


class NsysConverter:

    def __init__(self, file_path: str, sheet_name: str) -> None:
        self.file_path = file_path
        self.sheet_name = sheet_name

    def convert(self, column_name: str) -> None:
        print(f'File: {self.file_path}, Sheet: {self.sheet_name}, Column: {column_name}')
        file = pd.read_excel(self.file_path, sheet_name=self.sheet_name, header=0)
        col_edited = []
        for data in file[column_name]:
            col_edited.append(self._convert_num(data))
        col = f'{column_name}_edited'
        file[col] = pd.DataFrame(col_edited)
        file.to_excel(self.file_path, index=False)

    def _convert_num(self, data: str) -> float:
        value, unit = data.split()
        value = float(value)
        if unit == 's':
            return value
        elif unit == 'ms':
            return value * 0.001
        elif unit == 'μs':
            return value * 0.000001
        else:
            raise ValueError(f'Wrong value {value}!')


if __name__ == '__main__':
    from argparse import ArgumentDefaultsHelpFormatter, ArgumentParser
    parser = ArgumentParser(formatter_class=ArgumentDefaultsHelpFormatter)
    parser.add_argument('--file_path', type=str, required=True, help='File path of Nsyight Systems')
    parser.add_argument('--sheet_name', type=str, default='Sheet1', help='Sheet name of result file')
    parser.add_argument('--column_name', type=str, required=True, help='Target column name for convert')
    args = parser.parse_args()
    nc = NsysConverter(args.file_path, args.sheet_name)
    nc.convert(args.column_name)
