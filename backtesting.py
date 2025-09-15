import pandas as pd
pd.options.mode.chained_assignment = None

import openpyxl
from openpyxl import Workbook

import time

col_AssetSym = 'AssetSym' #
col_HldDys   = 'HoldDaysMax' #

col_Open     = 'Open' #
col_AdCl     = 'AdjClose' #

col_MAWin    = 'MAWin' #
col_MA200    = 'MA200'

col_PrChg    = 'PriceChange'
col_UpMv     = 'UpMove'
col_DnMv     = 'DownMove'

col_EMASpn   = 'EMASpan' #
col_EMAUp    = 'EMAvgUp'
col_EMADn    = 'EMAvgDown'

col_RS       = 'RS'
col_RSI      = 'RSI'

col_RSIMin   = 'RSIMin' #
col_RSIMax   = 'RSIMax' #

col_Buy      = 'Buy'
col_Sell     = 'Sell'

def RSIcalc(df_1, asset_sym, hld_dys, MA_win, EMA_spn, RSI_min, RSI_max):
    """Calculates the Relative Strength Index (RSI) and generates buy/sell signals.

    Args:
        df_1 (pd.DataFrame): The input DataFrame with stock data. Must contain
            'AdjClose' and 'Open' columns.
        asset_sym (str): The symbol of the asset.
        hld_dys (int): The maximum number of days to hold the asset.
        MA_win (int): The window for the moving average calculation.
        EMA_spn (int): The span for the exponential moving average calculation.
        RSI_min (int): The minimum RSI value for a buy signal.
        RSI_max (int): The maximum RSI value for a sell signal.

    Returns:
        pd.DataFrame: The DataFrame with the calculated RSI and buy/sell signals.
    """
    df_1[col_AssetSym] = asset_sym
    df_1[col_HldDys] = hld_dys
    df_1[col_MAWin] = MA_win
    df_1[col_EMASpn] = EMA_spn
    df_1[col_RSIMin] = RSI_min
    df_1[col_RSIMax] = RSI_max

    df_1[col_MA200] = df_1[col_AdCl].rolling(window=MA_win).mean()

    df_1[col_PrChg] = df_1[col_AdCl].pct_change()
    df_1[col_UpMv] = df_1[col_PrChg].apply(lambda x: x if x > 0 else 0)
    df_1[col_DnMv] = df_1[col_PrChg].apply(lambda x: abs(x) if x < 0 else 0)

    df_1[col_EMAUp] = df_1[col_UpMv].ewm(span=EMA_spn).mean()
    df_1[col_EMADn] = df_1[col_DnMv].ewm(span=EMA_spn).mean()

    df_1 = df_1.dropna()

    df_1[col_RS] = df_1[col_EMAUp] / df_1[col_EMADn]
    df_1[col_RSI] = df_1[col_RS].apply(lambda x: 100 - (100 / (x + 1)))

    df_1.loc[(df_1[col_AdCl] > df_1[col_MA200]), col_Buy] = 'No'
    df_1.loc[(df_1[col_AdCl] == df_1[col_MA200]), col_Buy] = 'No'
    df_1.loc[(df_1[col_AdCl] < df_1[col_MA200]), col_Buy] = 'No'

    df_1.loc[(df_1[col_RSI] > RSI_min), col_Buy] = 'No'
    df_1.loc[(df_1[col_RSI] == RSI_min), col_Buy] = 'No'
    df_1.loc[(df_1[col_RSI] < RSI_min), col_Buy] = 'No'

    df_1.loc[(df_1[col_AdCl] < df_1[col_MA200]) | (df_1[col_RSI] < RSI_min),
             col_Buy] = 'No'

    df_1.loc[(df_1[col_AdCl] > df_1[col_MA200]) & (df_1[col_RSI] < RSI_min),
             col_Buy] = 'RSIMin_BuyPrcAbvMA2'

    df_1.loc[(df_1[col_AdCl] > df_1[col_MA200]), col_Sell] = 'No'
    df_1.loc[(df_1[col_AdCl] == df_1[col_MA200]), col_Sell] = 'No'
    df_1.loc[(df_1[col_AdCl] < df_1[col_MA200]), col_Sell] = 'No'

    df_1.loc[(df_1[col_RSI] > RSI_min), col_Sell] = 'No'
    df_1.loc[(df_1[col_RSI] == RSI_min), col_Sell] = 'No'
    df_1.loc[(df_1[col_RSI] < RSI_min), col_Sell] = 'No'

    df_1.loc[(df_1[col_RSI] > RSI_max), col_Sell] = 'No'
    df_1.loc[(df_1[col_RSI] == RSI_max), col_Sell] = 'No'
    df_1.loc[(df_1[col_RSI] < RSI_max), col_Sell] = 'No'

    df_1.loc[(df_1[col_AdCl] > df_1[col_MA200]) & (df_1[col_RSI] > RSI_max),
             col_Sell] = 'RSIMax_SellPrcAbvMA2'
    df_1.loc[(df_1[col_AdCl] < df_1[col_MA200]) & (df_1[col_RSI] > RSI_max),
             col_Sell] = 'RSIMax_SellPrcBlwMA2'

    return df_1











def get_signals_1(df_3, hld_dys):
    """Generates buy and sell signals based on the RSI strategy.

    Args:
        df_3 (pd.DataFrame): The input DataFrame with RSI and buy/sell signals.
        hld_dys (int): The maximum number of days to hold the asset.

    Returns:
        tuple: A tuple containing two lists:
            - BuyDtsLst (list): A list of buy dates.
            - SellDtsLst (list): A list of sell dates.
    """
    BuyDtsLst = []
    SellDtsLst = []

    EndDB = len(df_3)
    for j in range(EndDB):
        if j == (EndDB - 2):
            break
        if 'RSIMin_BuyPrcAbvMA2' in df_3[col_Buy].iloc[j]:
            BuyDtsLst.append(df_3.iloc[j + 1].name)
            for k in range(1, hld_dys + 1):
                StpSell = 'No'
                # if (df_3[col_Sell].iloc[j + k] == 'No') :
                if (df_3[col_Sell].iloc[j + k] == 'RSIMax_SellPrcAbvMA2'):
                    StpSell = 'Yes'
                elif (df_3[col_Sell].iloc[j + k] == 'RSIMax_SellPrcBlwMA2'):
                    StpSell = 'Yes'
                elif (j + k + 1) == (EndDB - 1):
                    StpSell = 'Yes'
                elif k == hld_dys:
                    StpSell = 'Yes'
                else:
                    a = 0

                if StpSell == 'Yes':
                    SellDtsLst.append(df_3.iloc[j + k + 1].name)
                    break
    return (BuyDtsLst, SellDtsLst)














def get_signals_2(df_3, asset_sym, hld_dys, MA_win, EMA_spn, RSI_min, RSI_max):
    """Generates detailed trading signals and returns them as lists.

    Args:
        df_3 (pd.DataFrame): The input DataFrame with RSI and buy/sell signals.
        asset_sym (str): The symbol of the asset.
        hld_dys (int): The maximum number of days to hold the asset.
        MA_win (int): The window for the moving average calculation.
        EMA_spn (int): The span for the exponential moving average calculation.
        RSI_min (int): The minimum RSI value for a buy signal.
        RSI_max (int): The maximum RSI value for a sell signal.

    Returns:
        tuple: A tuple containing lists of trading signals and metadata.
    """
    AssetSymLst = []
    HldDyLst = []

    MAwinLst = []
    EMAspnLst = []
    RSIminLst = []
    RSImaxLst = []

    RSI1Lst = []
    BuyDtsLst = []
    BuyAmtLst = []

    RSI2Lst = []
    SellDtsLst = []
    SellAmtLst = []

    SellNoteLst = []

    EndDB = len(df_3)
    for j in range(EndDB):
        # print(j,
        #       df_3['Buy'].iloc[j],
        #       EndDB)
        if j == (EndDB - 2):
            break
        if 'RSIMin_BuyPrcAbvMA2' in df_3[col_Buy].iloc[j]:
            AssetSymLst.append(asset_sym)
            HldDyLst.append(hld_dys)
            MAwinLst.append(MA_win)
            EMAspnLst.append(EMA_spn)
            RSIminLst.append(RSI_min)
            RSImaxLst.append(RSI_max)

            RSI1Lst.append(int(df_3[col_RSI].iloc[j]))
            BuyDtsLst.append(df_3.iloc[j + 1].name)
            BuyAmtLst.append(df_3[col_Open].iloc[j + 1])  # .Open)
            for k in range(1, hld_dys + 1):
                StpSell = 'No'
                # if (df_3[col_Sell].iloc[j + k] == 'No') :
                if (df_3[col_Sell].iloc[j + k] == 'RSIMax_SellPrcAbvMA2'):
                    StpSell = 'Yes'
                    if df_3[col_Open].iloc[j + 1] < df_3[col_Open].iloc[j + k + 1]:
                        SellNoteLst.append('RSIMax_SellPrcAbvMA2_Prft')
                    elif df_3[col_Open].iloc[j + 1] == df_3[col_Open].iloc[j + k + 1]:
                        SellNoteLst.append('RSIMax_SellPrcAbvMA2_Nil')
                    else:
                        SellNoteLst.append('RSIMax_SellPrcAbvMA2_Loss')
                elif (df_3[col_Sell].iloc[j + k] == 'RSIMax_SellPrcBlwMA2'):
                    StpSell = 'Yes'
                    if df_3[col_Open].iloc[j + 1] < df_3[col_Open].iloc[j + k + 1]:
                        SellNoteLst.append('RSIMax_SellPrcBlwMA2_Prft')
                    elif df_3[col_Open].iloc[j + 1] == df_3[col_Open].iloc[j + k + 1]:
                        SellNoteLst.append('RSIMax_SellPrcBlwMA2_Nil')
                    else:
                        SellNoteLst.append('RSIMax_SellPrcBlwMA2_Loss')
                elif (j + k + 1) == (EndDB - 1):
                    StpSell = 'Yes'
                    if df_3[col_Open].iloc[j + 1] < df_3[col_Open].iloc[j + k + 1]:
                        SellNoteLst.append('DBMax_Prft')
                    elif df_3[col_Open].iloc[j + 1] == df_3[col_Open].iloc[j + k + 1]:
                        SellNoteLst.append('DBMax_Nil')
                    else:
                        SellNoteLst.append('DBMax_Loss')
                elif k == hld_dys:
                    StpSell = 'Yes'
                    # to check sell price and MA200
                    if df_3[col_Open].iloc[j + 1] < df_3[col_Open].iloc[j + k + 1]:
                        SellNoteLst.append('HoldMax_Prft')
                    elif df_3[col_Open].iloc[j + 1] == df_3[col_Open].iloc[j + k + 1]:
                        SellNoteLst.append('HoldMax_Nil')
                    else:
                        SellNoteLst.append('HoldMax_Loss')
                else:
                    a = 0

                if StpSell == 'Yes':
                    RSI2Lst.append(int(df_3[col_RSI].iloc[j + k]))
                    SellDtsLst.append(df_3.iloc[j + k + 1].name)
                    SellAmtLst.append(df_3[col_Open].iloc[j + k + 1])  # .Open))
                    break
    return (AssetSymLst,
            HldDyLst,
            MAwinLst,
            EMAspnLst,
            RSIminLst,
            RSImaxLst,
            RSI1Lst,
            BuyDtsLst,
            BuyAmtLst,
            RSI2Lst,
            SellDtsLst,
            SellAmtLst,
            SellNoteLst)



def main():
    """Main function to run the backtesting script."""
    start_time_1 = time.strftime("%H:%M:%S", time.localtime())

    data_fl_path = '/content/drive/MyDrive/data/'
    data_fl_path = '/storage/emulated/0/000_Python_Practise/02_Sample/'
    df_data_file = 'Stock_Summary_ITC.xlsx'

    stk_summ_fl_nm_0 = data_fl_path + df_data_file
    # stk_summ_fl_nm_0 = df_data_file
    # print(stk_summ_fl_nm_0)

    df = pd.read_excel(stk_summ_fl_nm_0)
    # print(df)

    df.set_index('Date', inplace=True)
    # print(df.head())

    xa = 'Stock_Summary_ITC_1.xlsx'
    stk_summ_fl_nm_1 = data_fl_path + xa
    # print(stk_summ_fl_nm_1)

    xabc = 'Stock_Summary_ITC_2.xlsx'
    stk_summ_fl_nm_2 = data_fl_path + xabc
    # print(stk_summ_fl_nm_2)

    wb_1 = openpyxl.load_workbook(stk_summ_fl_nm_2)
    ws_4 = wb_1["Data_0"]

    assetsym = 'ITC.NS'

    tot_rows = 0

    len_0 = 0
    sum_0 = 0
    mean_0 = 0

    wins = 0
    winrate = 0

    sum_p0 = 0
    sum_n0 = 0

    sum_1 = 0
    mean_1 = 0

    for holddays in range(50, 0, -1):
        for MAwin in range(200, 0, -1):
            for spn in range(48, 0, -1):
                for rmn in range(1, 31, 1):
                    # for rmn in range(     31, 41,  1) :
                    # for rmn in range(     41, 51,  1) :
                    # for rmn in range(     51, 61,  1) :
                    # for rmn in range(     61, 71,  1) :
                    # for rmn in range(     71, 81,  1) :
                    # for rmn in range(     81, 91,  1) :
                    # for rmn in range(     91, 100,  1) :

                    # time.sleep(0.5)
                    for rmx in range(100, rmn, -1):
                        print('a', holddays, MAwin, spn, rmn, rmx)
                        df_2 = RSIcalc(df,
                                       assetsym,
                                       holddays,
                                       MAwin,
                                       spn,
                                       rmn,
                                       rmx)
                        # print(df_2.head(10))
                        # print(len(df_2))
                        # df_2.to_excel(stk_summ_fl_nm_1) # ,
                        #              index = False)

                        (buydL,
                         selldL) = get_signals_1(df_2,
                                                holddays)

                        # (AssetSymL,
                        #  HldDyL,

                        #  MAwinL,
                        #  EMAspnL,
                        #  RSIminL,
                        #  RSImaxL,

                        #  RSI1L,
                        #  buydL,
                        #  buyaL,

                        #  RSI2L,
                        #  selldL,
                        #  sellaL,

                        #  sellnL) = get_signals_2(df_2,
                        #                        assetsym,
                        #                        holddays,
                        #                        MAwin,
                        #                        spn,
                        #                        rmn,
                        #                        rmx)

                        print('b', len(buydL), len(selldL))
                        # print('c', sellnL)

                        prftamt = []
                        prftpcent = []
                        if len(buydL) > 0:
                            tot_rows += 1
                            prftamt = df_2.loc[selldL].Open.values - df_2.loc[buydL].Open.values
                            # print('d', len(prftamt), prftamt)
                            prftpcent = (prftamt / df_2.loc[buydL].Open.values) * 100
                            len_0 = len(prftamt)
                            sum_0 = sum(prftamt)
                            mean_0 = prftamt.mean()
                            wins = [k for k in prftamt if k > 0]
                            winrate = (len(wins) / len_0) * 100
                            sum_p0 = sum(wins)
                            sum_n0 = sum_0 - sum_p0
                            if sum_1 == 0:
                                sum_1 = sum_0
                                mean_1 = mean_0
                            if sum_1 < sum_0:
                                sum_1 = sum_0
                                mean_1 = mean_0

                            if winrate < 50:
                                ws_4 = wb_1["Data_0"]
                            if winrate >= 50 and winrate < 60:
                                ws_4 = wb_1["Data_1"]
                            if winrate >= 60 and winrate < 70:
                                ws_4 = wb_1["Data_2"]
                            if winrate >= 70 and winrate < 80:
                                ws_4 = wb_1["Data_3"]
                            if winrate >= 80 and winrate < 90:
                                ws_4 = wb_1["Data_4"]
                            if winrate >= 90 and winrate < 100:
                                ws_4 = wb_1["Data_5"]
                            if winrate == 100:
                                ws_4 = wb_1["Data_6"]
                            ws_4.append([assetsym,
                                         holddays,
                                         MAwin,
                                         spn,
                                         rmn,
                                         rmx,
                                         sum_0,
                                         len_0,
                                         mean_0,
                                         sum_p0,
                                         sum_n0,
                                         len(wins),
                                         winrate,
                                         sum_1,
                                         mean_1])
                            print(
                                f'e {holddays} {MAwin} {spn} {rmn} {rmx} {len(buydL)} {sum_0:.2f} {mean_0:.2f} {winrate:.2f}')
                        # break
                    # break
                break
            break
        break

    print(1, start_time_1)
    print(2, time.strftime("%H:%M:%S", time.localtime()))
    print('Completed', tot_rows)
    wb_1.save(stk_summ_fl_nm_2)
    wb_1.close


if __name__ == "__main__":
    main()
