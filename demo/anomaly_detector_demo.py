import pandas as pd
s = pd.read_csv('D:/data/training.csv', index_col="Time", parse_dates=True, squeeze=True)
print(s)
from adtk.data import validate_series
s = validate_series(s)
from adtk.detector import ThresholdAD
threshold_ad = ThresholdAD(high=1300, low=900)
anomalies = threshold_ad.detect(s)
from adtk.visualization import plot
plot(s, anomaly=anomalies, ts_linewidth=1, ts_markersize=3, anomaly_markersize=5, anomaly_color='red', anomaly_tag="marker")
