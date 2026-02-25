<?php

$git_projects     = json_decode( file_get_contents( __DIR__ . '/git-projects.json' ), true );
$git_projects_map = array();
$projects         = array();

foreach ( $git_projects as $project ) {
    $git_projects_map[ $project['name'] ] = $project;
}

$show_projects = array(
    'flexole' => '#4d779f',
    'statik2' => '#c47a2d',
    'byrate'  => '#3f8857',
    'kriptun' => '#289487',
    'stego'   => '#b46b82',
    'dnsrv'   => '#95a536',
    'uconn'   => '#284f94',
    'logs'    => '#db6464',
    'gits'    => '#951e72',
    '?'       => '#423658',
);

foreach ( $show_projects as $name => $color ) {
    if ( $name === '?' ) {
        $projects[] = array(
            'icon'  => '?',
            'title' => '---- -- ---- --',
            'desc'  => '---- -- ---- -- ---- -- ---- --',
            'url'   => '?',
            'color' => $color,
        );

        continue;
    }

    $project = $git_projects_map[ $name ];

    $projects[] = array(
        'icon'  => strtoupper( substr( $project['name'], 0, 1 ) ),
        'title' => ucfirst( $project['name'] ),
        'desc'  => $project['description'],
        'url'   => $project['html_url'],
        'color' => $color,
    );
}

return $projects;