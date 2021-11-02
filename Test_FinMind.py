from FinMind.data import DataLoader


api = DataLoader()


stock_array = [1301, 1308, 1711, 1712, 1773, 2340, 2351, 2356, 2476, 2881, 2882, 2883, 2887, 2890, 2892, 2915, 3017, 3033, 3035, 3036,
               3141, 3189, 3491, 3527, 3551, 3587, 3675, 4721, 4739, 4927, 5347, 5351, 5434, 6104, 6138, 6170, 6411, 6509, 6592, 8069, 8155, 9945]


all_stock = api.taiwan_stock_info()


print(all_stock)

for n in range(len(stock_array)):
    stock = api.taiwan_stock_daily(
        stock_id=stock_array[n],
        start_date='2021-05-01',
        end_date='2021-11-02'
    )

    stock = stock.loc[:, ['stock_id', 'date', 'close']]
    list_price = stock["close"].to_list()

    max_price = max(list_price)
    min_price = min(list_price)

    # if list_price[-1] >= ((max_price - min_price)*4/5)+min_price:
    #    print('stock_id = ', stock_array[n])
