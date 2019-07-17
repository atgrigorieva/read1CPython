import win32com.client
import psycopg2
from psycopg2 import sql
import sched
import time
import threading
import logging
import re

logger = logging.getLogger()
logger.setLevel(logging.DEBUG)
handler = logging.FileHandler('viber bot.log')
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)

try:
    CON_STR = 'Srvr="127.0.0.1";Ref="trade 8.2";Usr="user.";Pwd="pass"'

    v83 = win32com.client.Dispatch("V83.COMConnector").Connect(CON_STR)
except Exception as e:
    logger.error('Failed to upload to ftp: ' + str(e))

#input("Press Enter to continue...")


from datetime import date, datetime, timedelta


def StockPrice(date_):
    print("Читаю акционную цену")
    query = '''ВЫБРАТЬ
    	ЦеныНоменклатурыСрезПоследних.Цена,
    	ЦеныНоменклатурыСрезПоследних.Номенклатура.Артикул КАК Артикул

    ИЗ
    	РегистрСведений.ЦеныНоменклатуры.СрезПоследних(ДАТАВРЕМЯ(%s, %s, %s, %s, %s, %s), ) КАК ЦеныНоменклатурыСрезПоследних
    ГДЕ
    	ЦеныНоменклатурыСрезПоследних.ТипЦен.Код = "УТ0000007"
    	И ЦеныНоменклатурыСрезПоследних.Номенклатура.ЭтоГруппа = 0
    	И ЦеныНоменклатурыСрезПоследних.Номенклатура.рв_ИДСайта <> "0"
    	И ЦеныНоменклатурыСрезПоследних.Номенклатура.рв_ИДСайта <> ""''' % (
        date_.year, date_.month, date_.day, date_.hour, date_.minute, date_.second)

    queryFlowFunds = v83.NewObject("Query", query)
    selectionPrice = queryFlowFunds.Execute().Choose()
    queryFlowFunds = None

    conn = psycopg2.connect(dbname='name_Bd', user='user_name',
                                password='password', host='#.#.#.#', port=5433)
    with conn.cursor() as cursor:
        while selectionPrice.Next():
            cursor.execute('Select id_product From products WHERE articul = %s', [selectionPrice.Артикул])

            output = cursor.fetchone()
            if output is not None:
                cursor.execute('Select id_price From prices Where id_product = %s and date_ = %s',
                               [output, date(date_.year, date_.month, date_.day)])

                outputPrice = cursor.fetchone()
                cursor.execute('UPDATE prices SET stock_price = %s WHERE id_price = %s',
                                                   [selectionPrice.Цена, outputPrice])
                conn.commit()


def PriceWrite(date_):
    print("Читаю Цены")
    query = '''ВЫБРАТЬ
	ЦеныНоменклатурыСрезПоследних.Цена,
	ЦеныНоменклатурыСрезПоследних.Номенклатура.Артикул КАК Артикул

ИЗ
	РегистрСведений.ЦеныНоменклатуры.СрезПоследних(ДАТАВРЕМЯ(%s, %s, %s, %s, %s, %s), ) КАК ЦеныНоменклатурыСрезПоследних
ГДЕ
	ЦеныНоменклатурыСрезПоследних.ТипЦен.Код = "000000001"
	И ЦеныНоменклатурыСрезПоследних.Номенклатура.ЭтоГруппа = 0
	И ЦеныНоменклатурыСрезПоследних.Номенклатура.рв_ИДСайта <> "0"
	И ЦеныНоменклатурыСрезПоследних.Номенклатура.рв_ИДСайта <> ""''' % (
    date_.year, date_.month, date_.day, date_.hour, date_.minute, date_.second)

    queryFlowFunds = v83.NewObject("Query", query)
    selectionPrice = queryFlowFunds.Execute().Choose()
    queryFlowFunds = None

    conn = psycopg2.connect(dbname='name_Bd', user='user_name',
                                password='password', host='#.#.#.#', port=5433)
    with conn.cursor() as cursor:
        while selectionPrice.Next():
            cursor.execute('Select id_product From products WHERE articul = %s', [selectionPrice.Артикул])

            output = cursor.fetchone()
            if output is not None:
                cursor.execute('Select id_price From prices Where id_product = %s and date_ = %s',
                               [output, date(date_.year, date_.month, date_.day)])

                outputPrice = cursor.fetchone()

                if outputPrice is None:
                    cursor.execute('INSERT INTO prices(date_, id_product, price, time_) VALUES (%s, %s, %s, %s)',
                                   (date(date_.year, date_.month, date_.day), output, selectionPrice.Цена, date_.strftime('%Y-%m-%d %H:%M:%S.%f')))
                    conn.commit()

                    cursor.execute('Select id_price From prices Where id_product = %s and date_ = %s',
                                   [output, date(date_.year, date_.month, date_.day)])

                    outputPrice = cursor.fetchone()

                else:
                    cursor.execute('UPDATE prices SET price = %s, time_ = %s WHERE id_price = %s',
                                   [selectionPrice.Цена, date_.strftime('%Y-%m-%d %H:%M:%S.%f'), outputPrice])
                    conn.commit()

def QuantityWrite(date_):
    print("Читаю Склад")
    query = '''ВЫБРАТЬ
	ТоварыНаСкладахОстатки.КоличествоОстаток,
	ТоварыНаСкладахОстатки.Номенклатура.Артикул КАК Артикул
ИЗ
	РегистрНакопления.ТоварыНаСкладах.Остатки(ДАТАВРЕМЯ(%s, %s, %s, %s, %s, %s), ) КАК ТоварыНаСкладахОстатки
ГДЕ
	ТоварыНаСкладахОстатки.Склад.Код = "000000001"
	И ТоварыНаСкладахОстатки.Номенклатура.ЭтоГруппа = 0
	И ТоварыНаСкладахОстатки.Номенклатура.рв_ИДСайта <> "0"
	И ТоварыНаСкладахОстатки.Номенклатура.рв_ИДСайта <> ""''' % (
    date_.year, date_.month, date_.day, date_.hour, date_.minute, date_.second)

    queryFlowFunds = v83.NewObject("Query", query)
    selectionQuantity = queryFlowFunds.Execute().Choose()
    queryFlowFunds = None

    conn = psycopg2.connect(dbname='name_Bd', user='user_name',
                                password='password', host='#.#.#.#', port=5433)
    with conn.cursor() as cursor:

        while selectionQuantity.Next():
            cursor.execute('Select id_product From products WHERE articul = %s', [selectionQuantity.Артикул])

            output = cursor.fetchone()
            if output is not None:
                cursor.execute('Select id_quantity From quantities Where id_product = %s and date_ = %s',
                               [output, date(date_.year, date_.month, date_.day)])

                outputQuantity = cursor.fetchone()

                if outputQuantity is None:
                    cursor.execute('INSERT INTO quantities(date_, id_product, quantity, time_) VALUES (%s, %s, %s, %s)', (
                    date(date_.year, date_.month, date_.day), output, selectionQuantity.КоличествоОстаток, date_.strftime('%Y-%m-%d %H:%M:%S.%f')))
                    conn.commit()

                    cursor.execute('Select id_quantity From quantities Where id_product = %s and date_ = %s',
                                   [output, date(date_.year, date_.month, date_.day)])

                    outputQuantity = cursor.fetchone()



                else:
                    cursor.execute('UPDATE quantities SET quantity = %s, time_ = %s WHERE id_quantity = %s',
                                   [selectionQuantity.КоличествоОстаток, date_.strftime('%Y-%m-%d %H:%M:%S.%f'), outputQuantity])
                    conn.commit()



def main(date_):
    '''Выборка первая: проходим по всей номеклатуре, имеющей юлабовский ИД'''
    try:
        print(date_)
        query = '''ВЫБРАТЬ
                Номенклатура.Наименование,
                Номенклатура.Артикул,
                Номенклатура.рв_ИДСайта,
                Номенклатура.Ссылка,
	            Номенклатура.ВидНоменклатуры.Наименование КАК ВидНоменклатуры
            ИЗ
                Справочник.Номенклатура КАК Номенклатура
            ГДЕ
                Номенклатура.ЭтоГруппа = 0
                И Номенклатура.рв_ИДСайта <> "0"
                И Номенклатура.рв_ИДСайта <> ""'''

        queryFlowFunds = v83.NewObject("Query", query)
        selection = queryFlowFunds.Execute().Choose()

        queryFlowFunds = None
        conn = psycopg2.connect(dbname='name_Bd', user='user_name',
                                password='password', host='#.#.#.#', port=5433)
        with conn.cursor() as cursor:
            while selection.Next():

                # рв_ИДСайта = selection.рв_ИДСайта.rsplit("_", 2)
                рв_ИДСайта = selection.рв_ИДСайта


                cursor.execute('Select id_product From products WHERE articul = %s', [selection.Артикул])
                # print(selection.Артикул)
                outputProduct = cursor.fetchone()
                if outputProduct is None:

                    '''Если в БД еще нет товара с таким артикулом, то записываем его'''

                    cursor.execute('INSERT INTO products(articul, product, id_site) VALUES (%s, %s, %s)',
                                   (selection.Артикул, selection.Наименование, рв_ИДСайта))
                    conn.commit()
                else:

                    type_nom = re.sub(r'[\d]+', r'',selection.ВидНоменклатуры).strip()
                    #print(type_nom)
                    cursor.execute('Select id_category From categories Where name_category ~* %s', [type_nom.replace('. ', '')])
                    outputCategory = cursor.fetchone()
                    if outputCategory is not None:
                        #print(str(outputCategory[0]))
                        cursor.execute('Update products Set id_category = %s Where id_product = %s', [str(outputCategory[0]), str(outputProduct[0])])
                        conn.commit()


        '''Записываем цену номенклатуры на определенную дату'''
        PriceWrite(date_)

        '''Записываем количество товара на складе'''
        QuantityWrite(date_)

        '''Записываем акционную цену номенклатуры на определенную дату'''
        StockPrice(date_)

    except Exception as e:
        logger.error('Failed to upload to ftp: ' + str(e))

   # input("Press Enter to continue...")




if __name__ == '__main__':
    date_ = datetime.now()
    #date__round_microsec = round(date_.microsecond / 1000)  # number of zeroes to round
   # date__ = date_.replace(microsecond=date__round_microsec)
    try:
        main(date_)
    except Exception as e:
        logger.error('Failed to upload to ftp: ' + str(e))

    #input("Press Enter to continue...")
    '''scheduler = sched.scheduler(time.time, time.sleep)
    scheduler.enter(600, 1, main, (datetime.now(),))
    t = threading.Thread(target=scheduler.run)
    t.start()'''

    #main()
