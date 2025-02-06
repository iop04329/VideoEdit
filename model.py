from dataclasses import dataclass, field
from typing import Optional, List
from datetime import datetime
import pandas as pd
from pandas.core.series import Series

@dataclass
class 船隻資訊:
    設施名稱: Optional[str] = None
    巡查日期: Optional[datetime] = None
    巡查人員: Optional[str] = None
    船名: Optional[str] = None
    巡查類型: Optional[str] = None
    靠離時間: Optional[datetime] = None
    備註: Optional[str] = None

def 轉換為船隻資訊(data: List) -> 船隻資訊:
    try:
        return 船隻資訊(
            設施名稱=data[0] if isinstance(data[0], str) else None,
            巡查日期=data[1] if isinstance(data[1], datetime) else None,
            巡查人員=data[2] if isinstance(data[2], str) else None,
            船名=data[3] if isinstance(data[3], str) else None,
            巡查類型=data[4] if isinstance(data[4], str) else None,
            靠離時間=data[5] if isinstance(data[5], datetime) else None,
            備註=data[6] if isinstance(data[6], str) else None,
        )
    except (IndexError, ValueError, TypeError):
        return 船隻資訊()

# 轉換 DataFrame 為 船隻資訊 類型
def 轉換為船隻資訊_新格式(row:Series) -> 船隻資訊:
    try:
        return 船隻資訊(
            設施名稱=row['設施名稱'] if isinstance(row['設施名稱'], str) else None,
            巡查日期=row['巡查日期'] if isinstance(row['巡查日期'], pd.Timestamp) else None,
            巡查人員=row['巡查人員'] if isinstance(row['巡查人員'], str) else None,
            船名=row['船名'] if isinstance(row['船名'], str) else None,
            巡查類型=row['巡查類型'] if isinstance(row['巡查類型'], str) else None,
            靠離時間=row['靠離時間'] if isinstance(row['靠離時間'], pd.Timestamp) else None,
            備註=None  # 假設無備註欄位
        )
    except Exception as e:
        print(f"轉換錯誤: {e}")
        return 船隻資訊()