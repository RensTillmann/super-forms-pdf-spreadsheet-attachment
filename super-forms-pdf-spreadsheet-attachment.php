<?php
/**
 * Super Forms - PDF & Spreadsheet Attachment
 *
 * @package   Super Forms - PDF & Spreadsheet Attachment
 * @author    feeling4design
 * @link      http://codecanyon.net/item/super-forms-drag-drop-form-builder/13979866
 * @copyright 2017 by feeling4design
 *
 * @wordpress-plugin
 * Plugin Name: Super Forms - PDF & Spreadsheet Attachment
 * Plugin URI:  http://codecanyon.net/item/super-forms-drag-drop-form-builder/13979866
 * Description: Sends a PDF or Spreadsheet attachment based on a spreadsheet template with possibility to dynamically include form data with the use of {tags}
 * Version:     1.0.0
 * Author:      feeling4design
 * Author URI:  http://codecanyon.net/user/feeling4design
*/

require_once( __DIR__ . '/vendor/autoload.php' );

if ( ! defined( 'ABSPATH' ) ) {
    exit; // Exit if accessed directly
}

if(!class_exists('SUPER_PDF_XLSX_Attachment')) :


    /**
     * Main SUPER_PDF_XLSX_Attachment Class
     *
     * @class SUPER_PDF_XLSX_Attachment
     * @version 1.0.0
     */
    final class SUPER_PDF_XLSX_Attachment {
    
        
        /**
         * @var string
         *
         *  @since      1.0.0
        */
        public $version = '1.0.0';


        /**
         * @var string
         *
         *  @since      1.0.0
        */
        public $add_on_slug = 'pdf_xlsx_attachment';
        public $add_on_name = 'PDF & Spreadsheet Attachment';

        
        /**
         * @var SUPER_PDF_XLSX_Attachment The single instance of the class
         *
         *  @since      1.0.0
        */
        protected static $_instance = null;

        
        /**
         * Main SUPER_PDF_XLSX_Attachment Instance
         *
         * Ensures only one instance of SUPER_PDF_XLSX_Attachment is loaded or can be loaded.
         *
         * @static
         * @see SUPER_PDF_XLSX_Attachment()
         * @return SUPER_PDF_XLSX_Attachment - Main instance
         *
         *  @since      1.0.0
        */
        public static function instance() {
            if(is_null( self::$_instance)){
                self::$_instance = new self();
            }
            return self::$_instance;
        }

        
        /**
         * SUPER_PDF_XLSX_Attachment Constructor.
         *
         *  @since      1.0.0
        */
        public function __construct(){
            $this->init_hooks();
            do_action('super_pdf_xlsx_attachment_loaded');
        }

        
        /**
         * Define constant if not already set
         *
         * @param  string $name
         * @param  string|bool $value
         *
         *  @since      1.0.0
        */
        private function define($name, $value){
            if(!defined($name)){
                define($name, $value);
            }
        }

        
        /**
         * What type of request is this?
         *
         * string $type ajax, frontend or admin
         * @return bool
         *
         *  @since      1.0.0
        */
        private function is_request($type){
            switch ($type){
                case 'admin' :
                    return is_admin();
                case 'ajax' :
                    return defined( 'DOING_AJAX' );
                case 'cron' :
                    return defined( 'DOING_CRON' );
                case 'frontend' :
                    return (!is_admin() || defined('DOING_AJAX')) && ! defined('DOING_CRON');
            }
        }

        
        /**
         * Hook into actions and filters
         *
         *  @since      1.0.0
        */
        private function init_hooks() {
            
            register_deactivation_hook( __FILE__, array( $this, 'deactivate' ) );
            
            // Filters since 1.0.0
            add_filter( 'super_after_activation_message_filter', array( $this, 'activation_message' ), 10, 2 );

            if ( $this->is_request( 'admin' ) ) {
                
                // Filters since 1.0.0
                add_filter( 'super_settings_after_smtp_server_filter', array( $this, 'add_settings' ), 10, 2 );
                add_filter( 'super_settings_end_filter', array( $this, 'activation' ), 100, 2 );

                // Actions since 1.0.0
                add_action( 'init', array( $this, 'update_plugin' ) );
                add_action( 'all_admin_notices', array( $this, 'display_activation_msg' ) );

            }
            
            if ( $this->is_request( 'ajax' ) ) {

                // Actions since 1.0.0
                add_action( 'super_before_sending_email_attachments_filter', array( $this, 'add_pdf_xlsx_admin_attachment' ), 10, 2 );
                add_action( 'super_before_sending_email_confirm_attachments_filter', array( $this, 'add_pdf_xlsx_confirm_attachment' ), 10, 2 );

            }
            
        }


        /**
         * Display activation message for automatic updates
         *
         *  @since      1.8.0
        */
        public function display_activation_msg() {
            if( !class_exists('SUPER_Forms') ) {
                echo '<div class="notice notice-error">'; // notice-success
                    echo '<p>';
                    echo sprintf( 
                        __( '%sPlease note:%s You must install and activate %4$s%1$sSuper Forms%2$s%5$s in order to be able to use %1$s%s%2$s!', 'super_forms' ), 
                        '<strong>', 
                        '</strong>', 
                        'Super Forms - ' . $this->add_on_name, 
                        '<a target="_blank" href="https://codecanyon.net/item/super-forms-drag-drop-form-builder/13979866">', 
                        '</a>' 
                    );
                    echo '</p>';
                echo '</div>';
            }
            $sac = get_option( 'sac_' . $this->add_on_slug, 0 );
            if( $sac!=1 ) {
                echo '<div class="notice notice-error">'; // notice-success
                    echo '<p>';
                    echo sprintf( __( '%sPlease note:%s You are missing out on important updates for %s! Please %sactivate your copy%s to receive automatic updates.', 'super_forms' ), '<strong>', '</strong>', 'Super Forms - ' . $this->add_on_name, '<a href="' . admin_url() . 'admin.php?page=super_settings#activate">', '</a>' );
                    echo '</p>';
                echo '</div>';
            }
        }


        /**
         * Automatically update plugin from the repository
         *
         *  @since      1.1.0
        */
        function update_plugin() {
            if( defined('SUPER_PLUGIN_DIR') ) {
                $sac = get_option( 'sac_' . $this->add_on_slug, 0 );
                if( $sac==1 ) {
                    require_once ( SUPER_PLUGIN_DIR . '/includes/admin/update-super-forms.php' );
                    $plugin_remote_path = 'http://f4d.nl/super-forms/';
                    $plugin_slug = plugin_basename( __FILE__ );
                    new SUPER_WP_AutoUpdate( $this->version, $plugin_remote_path, $plugin_slug, '', '', $this->add_on_slug );
                }
            }
        }

        
        /**
         * Add the activation under the "Activate" TAB
         * 
         * @since       1.0.0
        */
        public function activation($array, $data) {
            if (method_exists('SUPER_Forms','add_on_activation')) {
                return SUPER_Forms::add_on_activation($array, $this->add_on_slug, $this->add_on_name);
            }else{
                return $array;
            }
        }


        /**  
         *  Deactivate
         *
         *  Upon plugin deactivation delete activation
         *
         *  @since      1.0.0
         */
        public static function deactivate(){
            if (method_exists('SUPER_Forms','add_on_deactivate')) {
                SUPER_Forms::add_on_deactivate(SUPER_PDF_XLSX_Attachment()->add_on_slug);
            }
        }


        /**
         * Check license and show activation message
         * 
         * @since       1.0.0
        */
        public function activation_message( $activation_msg, $data ) {
            if (method_exists('SUPER_Forms','add_on_activation_message')) {
                $settings = $data['settings'];
                if( (isset($settings['pdf_xlsx_attachment_enable'])) && ($settings['pdf_xlsx_attachment_enable']=='true') ) {
                    return SUPER_Forms::add_on_activation_message($activation_msg, $this->add_on_slug, $this->add_on_name);
                }
            }
            return $activation_msg;
        }
        

        /**
         * Add attachment(s) to admin emails
         *
         *  @since      1.0.0
        */
        public static function add_pdf_xlsx_admin_attachment( $attachments, $data ) {
            if( (isset($data['settings']['pdf_xlsx_attachment_enable'])) && ($data['settings']['pdf_xlsx_attachment_enable']=='true') ) {
                if( !isset($data['settings']['pdf_xlsx_admin']) ) $data['settings']['pdf_xlsx_admin'] = '';
                if( $data['settings']['pdf_xlsx_admin']=='true' ) {
                    // Lets first find the template file and if not exists do nothing
                    if( !isset($data['settings']['pdf_xlsx_admin_template']) ) $data['settings']['pdf_xlsx_admin_template'] = '';
                    if( $data['settings']['pdf_xlsx_admin_template']!='' ) {
                        $templates = explode( ',', $data['settings']['pdf_xlsx_admin_template'] );
                        if( (!isset($data['settings']['pdf_xlsx_admin_name'])) || ($data['settings']['pdf_xlsx_admin_name']=='') ) {
                            $data['settings']['pdf_xlsx_admin_name'] = 'super-pdf-xlsx-attachment';
                        }
                        if( !isset($data['settings']['pdf_xlsx_admin_pdf']) ) {
                            $data['settings']['pdf_xlsx_admin_pdf'] = '';
                        }
                        $attachments = self::generate_attachment(
                            $attachments,
                            $data,
                            $templates,
                            $data['settings']['pdf_xlsx_admin_name'],
                            $data['settings']['pdf_xlsx_admin_pdf']
                        );
                    }
                }
            }
            return $attachments;
        }


        /**
         * Add attachment(s) to confirmation emails
         *
         *  @since      1.0.0
        */
        public static function add_pdf_xlsx_confirm_attachment( $attachments, $data ) {
            if( (isset($data['settings']['pdf_xlsx_attachment_enable'])) && ($data['settings']['pdf_xlsx_attachment_enable']=='true') ) {
                if( !isset($data['settings']['pdf_xlsx_confirm']) ) $data['settings']['pdf_xlsx_confirm'] = '';
                if( $data['settings']['pdf_xlsx_confirm']=='true' ) {
                    // Lets first find the template file and if not exists do nothing
                    if( $data['settings']['pdf_xlsx_confirm_template']!='' ) {
                        $templates = explode( ',', $data['settings']['pdf_xlsx_confirm_template'] );
                        if( (!isset($data['settings']['pdf_xlsx_confirm_name'])) || ($data['settings']['pdf_xlsx_confirm_name']=='') ) {
                            $data['settings']['pdf_xlsx_confirm_name'] = 'super-pdf-xlsx-attachment';
                        }
                        if( !isset($data['settings']['pdf_xlsx_confirm_pdf']) ) {
                            $data['settings']['pdf_xlsx_confirm_pdf'] = '';
                        }
                        $attachments = self::generate_attachment( 
                            $attachments, 
                            $data, 
                            $templates, 
                            $data['settings']['pdf_xlsx_confirm_name'], 
                            $data['settings']['pdf_xlsx_confirm_pdf'] 
                        );
                    }
                }
            }
            return $attachments;
        }

        /**
         * Generate attachment based on spreadsheet template
         *
         *  @since      1.0.0
        */
        public static function generate_attachment( $attachments, $data, $templates, $filename, $pdf='' ) {

            if( is_array($templates) ) {
                foreach( $templates as $k => $v ) {
                    $file = get_attached_file( $v );
                    if( $file ) {
                        $extension = '.xlsx';
                        if( $pdf=='true' ) {
                            $extension = '.pdf';
                            if( $data['settings']['pdf_xlsx_pdf_library']=='tcPDF' ) {
                                if( !class_exists( 'TCPDF' ) ) {
                                    require_once( 'lib/tcpdf/tcpdf.php' );
                                }
                            }
                        }
                        $file_location = '/uploads/php/files/' . sanitize_title_with_dashes($filename) . $extension;
                        $source = urldecode( SUPER_PLUGIN_DIR . $file_location );
                        if( file_exists( $source ) ) {
                            SUPER_Common::delete_file( $source );
                        }
                        error_reporting(E_ALL);
                        ini_set('display_errors', TRUE);
                        ini_set('display_startup_errors', TRUE);
                        date_default_timezone_set('Europe/London');
                        define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

                        if( !class_exists( 'PHPExcel_Settings' ) ) {
                            include_once( 'classes/PHPExcel/Settings.php' );
                        }
                        if( !class_exists( 'PHPExcel_IOFactory' ) ) {
                            include_once( 'classes/PHPExcel/IOFactory.php' );
                        }
                        if( !class_exists( 'PHPExcel' ) ) {
                            include_once( 'classes/PHPExcel.php' );
                        }

                        //  Create PDF
                        $excel = PHPExcel_IOFactory::createReader('Excel2007');
                        $excel = $excel->load($file); // Empty Sheet
                        if( $pdf ) {
                            $excel->getActiveSheet()->setShowGridLines(false);
                            $excel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(-1);
                        }
                        foreach ($excel->getWorksheetIterator() as $worksheet) {
                            $ws = $worksheet->getTitle();
                            foreach( $worksheet->getRowIterator() as $row ) {
                                $cellIterator = $row->getCellIterator();
                                //$cellIterator->setIterateOnlyExistingCells(true);
                                foreach ($cellIterator as $cell) {
                                    if( ($cell->getValue()==null) || ($cell->getValue()=='') ) {
                                        $cell->setValue(' ');
                                        continue;
                                    } 
                                    if (!preg_match("/\{(.*?)\}/", $cell->getValue())) continue;
                                    $new_value = SUPER_Common::email_tags( $cell->getValue(), $data['data'], $data['settings'] );
                                    $cell->setValue($new_value);
                                }
                            }
                        }
                        if( $pdf ) {
                            if( !isset($data['settings']['pdf_xlsx_pdf_library']) ) {
                                $data['settings']['pdf_xlsx_pdf_library'] = 'tcPDF';
                            }
                            $objWriter = PHPExcel_IOFactory::createWriter($excel, 'PDF_' . $data['settings']['pdf_xlsx_pdf_library']);
                        }else{
                            $objWriter = PHPExcel_IOFactory::createWriter($excel, 'Excel2007');
                        }
                        $objWriter->save($source);
                        $attachments[$filename . $extension] = SUPER_PLUGIN_FILE . $file_location;
                    }
                }
            }
            return $attachments;

        }


        /**
         * Formats a line (passed as a fields  array) as CSV and returns the CSV as a string.
         *
         *  @since      1.0.0
        */
        public static function array_to_csv( array &$fields, $delimiter = ';', $enclosure = '"', $encloseAll = false, $nullToMysqlNull = false ) {
            $delimiter_esc = preg_quote($delimiter, '/');
            $enclosure_esc = preg_quote($enclosure, '/');
            $output = array();
            foreach ( $fields as $field ) {
                if ($field === null && $nullToMysqlNull) {
                    $output[] = 'NULL';
                    continue;
                }
                if ( $encloseAll || preg_match( "/(?:${delimiter_esc}|${enclosure_esc}|\s)/", $field ) ) {
                    $output[] = $enclosure . str_replace($enclosure, $enclosure . $enclosure, $field) . $enclosure;
                }
                else {
                    $output[] = $field;
                }
            }
            return implode( $delimiter, $output );
        }


        /**
         * Hook into settings and add PDF & Spreadsheet settings
         *
         *  @since      1.0.0
        */
        public static function add_settings( $array, $settings ) {

            $array['pdf_xlsx_attachment'] = array(        
                'hidden' => 'settings',
                'name' => __( 'PDF & Spreadsheet Attachment', 'super-forms' ),
                'label' => __( 'PDF & Spreadsheet Attachment', 'super-forms' ),
                'fields' => array(
                    'pdf_xlsx_attachment_enable' => array(
                        'desc' => __( 'This will attach a PDF or Spreadsheet file to the admin email', 'super-forms' ), 
                        'default' => SUPER_Settings::get_value( 0, 'pdf_xlsx_attachment_enable', $settings['settings'], '' ),
                        'type' => 'checkbox',
                        'values' => array(
                            'true' => __( 'Attach PDF(s) or Spreadsheet(s)', 'super-forms' ),
                        ),
                        'filter' => true
                    ),

                    // PDF rendering library
                    'pdf_xlsx_pdf_library' => array(
                        'name' => __( 'Choose your preferred PDF rendering library', 'super-forms' ), 
                        'label' => __( 'The different libraries have different strengths and weaknesses. Some generate better formatted output than others, some are faster or use less memory than others, while some generate smaller .pdf files. It is the developers choice which one they wish to use, appropriate to their own circumstances.', 'super-forms' ), 
                        'default' => SUPER_Settings::get_value( 0, 'pdf_xlsx_pdf_library', $settings['settings'], 'tcPDF' ),
                        'type' => 'select',
                        'values' => array(
                            'tcPDF' => __( 'TCPDF library (default)', 'super-forms' ),
                            'mPDF' => __( 'mPDF library ', 'super-forms' ),
                            'DomPDF' => __( 'Dompdf library', 'super-forms' ),
                        ),
                        'filter'=>true,
                        'parent'=>'pdf_xlsx_attachment_enable',
                        'filter_value'=>'true'
                    ),

                    // Admin email attachment(s)
                    'pdf_xlsx_admin' => array(
                        'desc' => __( 'Send attachment for admin emails', 'super-forms' ), 
                        'default' => SUPER_Settings::get_value( 0, 'pdf_xlsx_admin', $settings['settings'], '' ),
                        'type' => 'checkbox',
                        'values' => array(
                            'true' => __( 'Enable attachment(s) for admin emails', 'super-forms' ),
                        ),
                        'filter'=>true,
                        'parent'=>'pdf_xlsx_attachment_enable',
                        'filter_value'=>'true'
                    ),
                    'pdf_xlsx_admin_template' => array(
                        'name' => __( 'Select template file(s) for admin emails (required)', 'super-forms' ),
                        'label' => __( 'You are allowed to use {tags} inside your spreadsheets', 'super-forms' ),
                        'default' => SUPER_Settings::get_value( 0, 'pdf_xlsx_admin_template', $settings['settings'], '' ),
                        'type' => 'file',
                        'multiple' => 'true',
                        'file_type'=>'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        'filter'=>true,
                        'parent'=>'pdf_xlsx_admin',
                        'filter_value'=>'true'
                    ),
                    'pdf_xlsx_admin_pdf' => array(
                        'desc' => __( 'This will send a PDF instead of Spreadsheet attachment for admin emails', 'super-forms' ), 
                        'label' => __( 'Please note that PDF file format has some limits regarding to styling cells, number formatting, ...', 'super-forms' ), 
                        'default' => SUPER_Settings::get_value( 0, 'pdf_xlsx_admin_pdf', $settings['settings'], '' ),
                        'type' => 'checkbox',
                        'values' => array(
                            'true' => __( 'Attach as PDF file instead of Spreadsheet file for admin emails', 'super-forms' ),
                        ),
                        'filter'=>true,
                        'parent'=>'pdf_xlsx_admin',
                        'filter_value'=>'true'
                    ),
                    'pdf_xlsx_admin_name' => array(
                        'name'=> __( 'The admin filename(s) of the attachment(s)', 'super-forms' ),
                        'label' => __( 'If you have multiple template files, seperate each filename by pipes: |', 'super-forms' ),
                        'placeholder' => __( 'filename1|filename2|filename3', 'super-forms' ),
                        'default'=> SUPER_Settings::get_value( 0, 'pdf_xlsx_admin_name', $settings['settings'], 'super-pdf-xlsx-attachment' ),
                        'filter'=>true,
                        'parent'=>'pdf_xlsx_admin',
                        'filter_value'=>'true'
                    ),

                    // Confirmation email attachment(s)
                    'pdf_xlsx_confirm' => array(
                        'desc' => __( 'Send attachment for confirmation emails', 'super-forms' ), 
                        'default' => SUPER_Settings::get_value( 0, 'pdf_xlsx_confirm', $settings['settings'], '' ),
                        'type' => 'checkbox',
                        'values' => array(
                            'true' => __( 'Enable attachment(s) for confirmation emails', 'super-forms' ),
                        ),
                        'filter'=>true,
                        'parent'=>'pdf_xlsx_attachment_enable',
                        'filter_value'=>'true'
                    ),
                    'pdf_xlsx_confirm_template' => array(
                        'name' => __( 'Select template file(s) for confirmation emails (required)', 'super-forms' ),
                        'label' => __( 'You are allowed to use {tags} inside your spreadsheets', 'super-forms' ),
                        'default' => SUPER_Settings::get_value( 0, 'pdf_xlsx_confirm_template', $settings['settings'], '' ),
                        'type' => 'file',
                        'multiple' => 'true',
                        'file_type'=>'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        'filter'=>true,
                        'parent'=>'pdf_xlsx_confirm',
                        'filter_value'=>'true'
                    ),
                    'pdf_xlsx_confirm_pdf' => array(
                        'desc' => __( 'This will send a PDF instead of Spreadsheet attachment for confirmation emails', 'super-forms' ), 
                        'label' => __( 'Please note that PDF file format has some limits regarding to styling cells, number formatting, ...', 'super-forms' ), 
                        'default' => SUPER_Settings::get_value( 0, 'pdf_xlsx_confirm_pdf', $settings['settings'], '' ),
                        'type' => 'checkbox',
                        'values' => array(
                            'true' => __( 'Attach as PDF file instead of Spreadsheet file for confirmation emails', 'super-forms' ),
                        ),
                        'filter'=>true,
                        'parent'=>'pdf_xlsx_confirm',
                        'filter_value'=>'true'
                    ),
                    'pdf_xlsx_confirm_name' => array(
                        'name'=> __( 'The confirmation filename(s) of the attachment(s)', 'super-forms' ),
                        'label' => __( 'If you have multiple template files, seperate each filename by pipes: |', 'super-forms' ),
                        'placeholder' => __( 'filename1|filename2|filename3', 'super-forms' ),
                        'default'=> SUPER_Settings::get_value( 0, 'pdf_xlsx_confirm_name', $settings['settings'], 'super-pdf-xlsx-attachment' ),
                        'filter'=>true,
                        'parent'=>'pdf_xlsx_confirm',
                        'filter_value'=>'true'
                    ),
                  
                )
            );
            return $array;
        }



    }
        
endif;


/**
 * Returns the main instance of SUPER_PDF_XLSX_Attachment to prevent the need to use globals.
 *
 * @return SUPER_PDF_XLSX_Attachment
 */
function SUPER_PDF_XLSX_Attachment() {
    return SUPER_PDF_XLSX_Attachment::instance();
}


// Global for backwards compatibility.
$GLOBALS['SUPER_PDF_XLSX_Attachment'] = SUPER_PDF_XLSX_Attachment();