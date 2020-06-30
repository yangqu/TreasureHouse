import pandas as pd


def main():

    ratio = 0.2
    push_times_up = 7
    push_times_bottom = 4
    playlog_times_up = 1200
    playlog_times_bottom = 600

    dir = r'D:\data'
    file = r'output.txt'
    io = dir + '\\' + file
    out_good = dir + '\\out_good.txt'
    out_bad = dir + '\\out_bad.txt'
    out_common = dir + '\\out_common.txt'
    source = pd.read_csv(io, error_bad_lines=False, sep='\t')
    source_extract = source[['locationid', 'stat1', 'stat2', 'stat3', 'num']]
    source_extract.fillna(0, inplace=True)
    good_device_num = int(round(source_extract['locationid'].count() * ratio))
    source_good = source_extract\
        .query('stat1 > {push_times_up} & stat2 > {playlog_times_up}'
               .format(push_times_up=push_times_up, playlog_times_up=playlog_times_up)).head(good_device_num)
    source_bad = source_extract\
        .query('stat1 < {push_times_bottom} & stat2 < {playlog_times_bottom}'
               .format(push_times_bottom=push_times_bottom, playlog_times_bottom=playlog_times_bottom)).tail(good_device_num)
    source_extract = source_extract.append(source_good)
    source_extract = source_extract.append(source_bad)
    source_common = source_extract.drop_duplicates(subset=['locationid', 'stat1', 'stat2', 'stat3', 'num'], keep=False)
    source_common.to_csv(out_common, sep='\t')
    source_good.to_csv(out_good, sep='\t')
    source_bad.to_csv(out_bad, sep='\t')
    print(source_common)


if __name__ == '__main__':
    main()
