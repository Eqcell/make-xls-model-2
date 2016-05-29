План:

1. исправляем ошибки в одном [файле тестов](https://github.com/epogrebnyak/make-xls-model/issues/34) 
  - если py.test проходит мерджим бранч c тестами в мастер
  - завершаем обсуждение списка тестов в <https://github.com/epogrebnyak/make-xls-model/issues/25>
2. утверждаем [стуктуру классов](https://github.com/epogrebnyak/make-xls-model/issues/29) + пишем классы которые проходят те же тесты
3. смотрим что можно сделать по поводу [объединения 'data' и 'controls'][dc] как отдельная ветка
  - нужен набор правил по которым разбираются строки данных и параметров
  - программа может раскидывать шит 'dataset' на шиты 'data' и 'controls' 
  - демонстрация результататов на bank_sector.xls
4. прочие [issues](https://github.com/epogrebnyak/make-xls-model/issues)
5. новые усовершенствования
 - без xlwings
 - изменения, необходимые для построения модели на нескольких шитах
 - резервное копирование файла под другим именем перед запуском программы как защита от перезаписи
 - что-то еще глобальное?

[dc]:https://github.com/epogrebnyak/make-xls-model/issues?q=is%3Aissue+is%3Aopen+label%3Aenhancement



