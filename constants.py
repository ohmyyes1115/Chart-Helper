
# qd120_[x]_qd100   並非固定 1=285889258 2=285889260，如果個案只有CA03，那 1=285889260
from enum import Enum

TAKE_CARE_HOME_PAGE_URL = r'https://csms2.sfaa.gov.tw/lcms/saTree/treeIndex?saPortalId=278266377'


class HomeTherapyCompany(Enum):
    CC = 'CC'
    LC = 'LC'


CLIENT_TYPE_ID_CA01 = r'285889258'
CLIENT_TYPE_ID_CA03 = r'285889260'

CC_ACCOUNT = 'CCO0168'
CC_PASSWORD = CC_ACCOUNT
LC_ACCOUNT = 'LC035'
LC_PASSWORD = '123456'
