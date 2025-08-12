<?php

function data($key) {
    $steps = explode('.', $key);
    $file  = __DIR__ . '/' . implode("/", $steps) . '.php';

    if ( ! file_exists( $file ) ) {
        return null;
    }

    return include( $file );
}