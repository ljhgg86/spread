<?php
    include 'conf.php';
    // ini_set('display_errors', 0); 

    $mysqli = new mysqli($conf['host'],$conf['username'],$conf['password'],$conf['database']);
    
    /* check connection */
    if (mysqli_connect_errno()) {
        printf("Connect failed: %s\n", mysqli_connect_error());
        exit();
    }
    $mysqli->set_charset("utf8");
?>