<?php

require_once 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class pwexcelstatic extends Module
{
    public function __construct()
    {
        $this->name = get_class($this);
        $this->version = '0.1.0';
        $this->author = 'PrestaWeb';

        $this->bootstrap = true;
        parent::__construct();

        $this->displayName = $this->l('Excel Static');
        $this->description = $this->l('Excel Static');

        $this->need_instance = 0;
        $this->ps_versions_compliancy = array('min' => '1.5.0.0', 'max' => _PS_VERSION_);
    }

    protected function renderForm()
    {
        if (Tools::isSubmit('submitGenerate')) {
            $this->postProcess();
        }

        $fields_form = array(
            'form' => array(
                'legend' => array(
                    'title' => 'Отчет',
                    'icon' => 'icon-cogs'
                ),
                'input' => array(
                    array(
                        'type' => 'date',
                        'label' => 'Дата начала:',
                        'name' => 'date_start',
                    ),
                    array(
                        'type' => 'date',
                        'label' => 'Дата конца:',
                        'name' => 'date_finish',
                    ),
                ),
                'submit' => array(
                    'title' => $this->l('Сгенерировать'),
                    'value' => 1
                )
            ),
        );
        
        $helper = new HelperForm();
        $helper->show_toolbar = false;
        $helper->submit_action = 'submitGenerate';
        $helper->token = Tools::getAdminTokenLite('AdminModules');
        
        return $helper->generateForm(array($fields_form));
    }

    protected function postProcess()
    {
        if (!Tools::getIsset('date_start') || !Tools::getIsset('date_finish')) {
            return;
        }

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setCellValue('A1', 'Дата');
        $sheet->setCellValue('B1', 'Артикул');
        $sheet->setCellValue('C1', 'Раздел');
        $sheet->setCellValue('D1', 'Поставщики');
        $sheet->setCellValue('E1', 'Название');
        $sheet->setCellValue('F1', 'Количество');
        $sheet->setCellValue('G1', 'Номер заказа');
        $sheet->setCellValue('H1', 'Цена закупки');
        $sheet->setCellValue('I1', 'Цена продажи');
        $sheet->setCellValue('J1', 'Статус заказа');

        $abc = array(
            'A', 'B', 'C', 'D', 'E',
            'F', 'G', 'H', 'I', 'J',
        );

        $dbQuery = new DbQuery();
        $dbQuery->select('o.id_order, o.date_add, o.reference, o.id_cart, osl.name');
        $dbQuery->from('orders', 'o');
        $dbQuery->innerJoin('order_state_lang', 'osl', 'o.current_state = osl.id_order_state');
        $dbQuery->where('o.date_add >= "' . Tools::getValue('date_start') . '"');
        $dbQuery->where('o.date_add <= "' . Tools::getValue('date_finish') . '"');
 
        $orders = DB::getInstance()->executeS($dbQuery);
        $index = 2;
        foreach ($orders as $i => $order) {
            $cart = new Cart($order['id_cart']);
            foreach ($cart->getProducts() as $cartProduct) {
                $product = new Product($cartProduct['id_product']);
                $sheet->setCellValue('A' . $index, $order['date_add']);
                $sheet->setCellValue('G' . $index, $order['reference']);
                $sheet->setCellValue('J' . $index, $order['name']);

                $sheet->setCellValue('B' . $index, $product->reference);
                $sheet->setCellValue('E' . $index, $product->name[1]);
                $sheet->setCellValue('F' . $index, $cartProduct['quantity']);
                $sheet->setCellValue('H' . $index, $product->wholesale_price);
                $sheet->setCellValue('I' . $index, $product->price);

                $category = new Category($product->id_category_default);
                $sheet->setCellValue('C' . $index, $category->name[1]);
                
                $manufacturer = new Manufacturer($product->id_manufacturer);
                $manufacturerName = !is_null($manufacturer->name) ? $manufacturer->name: '-';
                $sheet->setCellValue('D' . $index, $manufacturerName);
                
                $index++;
            }

        }

        $writer = new Xlsx($spreadsheet);
        $writer->save(__DIR__ . '/tmp.xlsx');
        header('Location: /modules/' . $this->name . '/tmp.xlsx');
    }

    public function install()
    {
        return parent::install();
    }

    public function uninstall()
    {
        return parent::uninstall();
    }

    public function getContent()
    {
        return $this->renderForm();
    }
}
