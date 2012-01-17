<?php

/**
 (c) DHTMLX Ltd, 2011
 Licensing: You allowed to use this component for free under GPL or you need to obtain Commercial/Enterprise license to use it in non-GPL project
 Contact: sales@dhtmlx.com
 **/



/**
 * @package Spreadsheet
 * @author DHTMLX LTD
 * @version 1.0
 */
/*
Plugin Name: Spreadsheet
Plugin URI: http://wordpress.org/extend/plugins/spreadsheet/
Description: This plugin allows you to quickly create an Excel-like, editable spreadsheet with basic cell formatting and math functions support.
Author: DHTMLX LTD
Version: 1.0
Author URI: http://dhtmlx.com
*/

require_once(ABSPATH.WPINC.'/pluggable.php');
require_once(WP_PLUGIN_DIR.'/dhtmlxspreadsheet/spreadsheet_common.php');

global $wpdb, $sh_cfg;

/*! initialize configs
 */
$sh_cfg = new SpreadsheetCfg();
$sh_cfg->set('prefix', $wpdb->prefix);
$sh_cfg->set('plugin', WP_PLUGIN_URL.'/dhtmlxspreadsheet/');
$sh_cfg->set('connector', WP_PLUGIN_URL.'/dhtmlxspreadsheet/spreadsheet_data.php');
$sh_cfg->set('sheet', 1);

register_activation_hook(__FILE__, 'sh_activate');
add_filter('the_content', 'sh_check');

function sh_check($content) {

	global $sh_cfg;
	$ver = phpversion();
	$ver_main = (int) substr($ver, 0, 1);
	if ( $ver_main < 5)
		return __('Installation error: Spreadsheet plugin requires PHP 5.x', 'spreadsheet');
	$content = preg_replace_callback($sh_cfg->get('pattern'), "sh_replace", $content);
	return $content;
}

function sh_replace($matches) {
	global $sh_cfg;
	$options = Array();
	if (isset($matches[1])) {
		$ms = Array();
		$matches[1] = html_entity_decode($matches[1]);
		preg_match_all($sh_cfg->get('pattern_param'), $matches[1], $ms);
		
		if (isset($ms[1])) {
			for ($i = 0; $i < count($ms[1]); $i++) {
				$options[trim($ms[1][$i])] = trim($ms[2][$i]);
			}
		}
	}

	$matches = explode(":",$matches[1]);
	foreach ($options as $name => $value) {
		$sh_cfg->set($name, $value);
	}
	if (!isset($options['height']))
		$sh_cfg->set('autoheight', 'true');

	$text = $sh_cfg->get_spreadsheet_client();
	return $text;
}


function sh_activate() {
	sh_load_dump();
}


function sh_init() {
	global $sh_cfg;
	return $sh_cfg->get_spreadsheet_client();
}


function sh_load_dump($drop = false) {
	global $wpdb, $sh_cfg;
	$query = "SELECT * FROM ".$wpdb->prefix."data";
	$result = $wpdb->query($query);
	if ($result == false) {
		$queries = $sh_cfg->get('queries');
		for ($i = 0; $i < count($queries); $i++) {
			$query = str_replace('#__', $wpdb->prefix, $queries[$i]);
			$wpdb->query($query);
		}
	}
}


?>