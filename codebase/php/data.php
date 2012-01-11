<?php

/**
 (c) DHTMLX Ltd, 2011
 Licensing: You allowed to use this component for free under GPL or you need to obtain Commercial/Enterprise license to use it in non-GPL project
 Contact: sales@dhtmlx.com
 **/




require("config.php");
require("grid_cell_connector.php");

$res = mysql_connect($db_host, $db_user, $db_pass);
mysql_select_db($db_name, $res);

$conn = new GridCellConnector($res, $db_prefix);
//$conn->enable_log();
$conn->render();

?>