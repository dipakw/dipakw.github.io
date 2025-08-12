<?php require_once __DIR__ . '/functions.php'; ?><!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Dipak Acharya | dipakw</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link
        href="https://fonts.googleapis.com/css2?family=Google+Sans+Code:ital,wght@0,300..800;1,300..800&family=Nunito:ital,wght@0,200..1000;1,200..1000&family=Raleway:ital,wght@0,100..900;1,100..900&display=swap"
        rel="stylesheet">
    <link rel="stylesheet" type="text/css" href="https://cdn.jsdelivr.net/npm/toastify-js/src/toastify.min.css">
    <link rel="stylesheet" href="/styles.css">
</head>

<body>

    <header></header>

    <main>
        <div class="sided">
            <div class="left">
                <div class="intro">
                    <div><img src="/img/dipak.jpg" alt="Dipak Acharya" class="person-avt"></div>
                    <div class="name-links">
                        <div class="name">Dipak Acharya</div>
                        <div>|</div>
                        <a href="https://github.com/dipakw" class="iconed-link" target="_blank">
                            <svg width="20px" height="20px" viewBox="0 0 20 20" version="1.1"
                                xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"
                                fill="currentColor">
                                <g stroke-width="0"></g>
                                <g stroke-linecap="round" stroke-linejoin="round"></g>
                                <g>
                                    <defs> </defs>
                                    <g stroke="none" stroke-width="1" fill="none" fill-rule="evenodd">
                                        <g transform="translate(-140.000000, -7559.000000)" fill="currentColor">
                                            <g transform="translate(56.000000, 160.000000)">
                                                <path
                                                    d="M94,7399 C99.523,7399 104,7403.59 104,7409.253 C104,7413.782 101.138,7417.624 97.167,7418.981 C96.66,7419.082 96.48,7418.762 96.48,7418.489 C96.48,7418.151 96.492,7417.047 96.492,7415.675 C96.492,7414.719 96.172,7414.095 95.813,7413.777 C98.04,7413.523 100.38,7412.656 100.38,7408.718 C100.38,7407.598 99.992,7406.684 99.35,7405.966 C99.454,7405.707 99.797,7404.664 99.252,7403.252 C99.252,7403.252 98.414,7402.977 96.505,7404.303 C95.706,7404.076 94.85,7403.962 94,7403.958 C93.15,7403.962 92.295,7404.076 91.497,7404.303 C89.586,7402.977 88.746,7403.252 88.746,7403.252 C88.203,7404.664 88.546,7405.707 88.649,7405.966 C88.01,7406.684 87.619,7407.598 87.619,7408.718 C87.619,7412.646 89.954,7413.526 92.175,7413.785 C91.889,7414.041 91.63,7414.493 91.54,7415.156 C90.97,7415.418 89.522,7415.871 88.63,7414.304 C88.63,7414.304 88.101,7413.319 87.097,7413.247 C87.097,7413.247 86.122,7413.234 87.029,7413.87 C87.029,7413.87 87.684,7414.185 88.139,7415.37 C88.139,7415.37 88.726,7417.2 91.508,7416.58 C91.513,7417.437 91.522,7418.245 91.522,7418.489 C91.522,7418.76 91.338,7419.077 90.839,7418.982 C86.865,7417.627 84,7413.783 84,7409.253 C84,7403.59 88.478,7399 94,7399">
                                                </path>
                                            </g>
                                        </g>
                                    </g>
                                </g>
                            </svg>

                            GitHub
                        </a>
                    </div>
                </div>
                <div class="projects">
                    <h2 class="head">Open Source Projects</h2>
                    <div class="inner">
                        <?php foreach ( data('data.general.projects') as $project ) : ?>
                        <div class="project">
                            <div class="icon" style="background-color: <?php echo $project['color']; ?>;">
                                <?php echo $project['icon']; ?>
                            </div>
                            <div>
                                <div class="title"><?php echo $project['title']; ?></div>
                                <div class="desc"><?php echo $project['desc']; ?></div>

                                <div class="links">
                                    <?php if ( isset( $project['url'] ) ) : ?>
                                        <?php if ( $project['url'] === '?' ) : ?>
                                            <a>----</a>
                                        <?php else : ?>
                                            <a href="<?php echo $project['url']; ?>" target="_blank">GitHub</a>
                                        <?php endif; ?>
                                    <?php endif; ?>
                                </div>
                            </div>
                        </div>
                        <?php endforeach; ?>
                    </div>
                </div>
            </div>
            <div class="right">
                <div class="iconed-header">
                    <svg width="52px" height="52px" viewBox="0 0 512 512" xmlns="http://www.w3.org/2000/svg"
                        fill="#28944c">
                        <g stroke-width="0"></g>
                        <g stroke-linecap="round" stroke-linejoin="round"></g>
                        <g>
                            <title>ionicons-v5-l</title>
                            <rect x="32" y="48" width="448" height="416" rx="48" ry="48"
                                style="fill:none;stroke:#28944c;stroke-linejoin:round;stroke-width:32px"></rect>
                            <polyline points="96 112 176 176 96 240"
                                style="fill:none;stroke:#28944c;stroke-linecap:round;stroke-linejoin:round;stroke-width:32px">
                            </polyline>
                            <line x1="192" y1="240" x2="256" y2="240"
                                style="fill:none;stroke:#28944c;stroke-linecap:round;stroke-linejoin:round;stroke-width:32px">
                            </line>
                        </g>
                    </svg>

                    <div>
                        <div class="title">Scripts</div>
                        <div class="desc">Let the scripts handle it while you take it easy.</div>
                    </div>
                </div>

                <div class="terminal">
                    <div class="terminal-head"></div>
                    <div class="terminal-body">
                        <div class="block">
                            <div class="line comment">
                                Download flexole - a secure reverse proxy.
                                <i class="linux"></i><i class="macos"></i><i class="gitbash"></i>
                            </div>
                            <div class="line command"><span class="cmd">curl</span> <span class="opts">-sL</span>
                                <span class="url">https://dipakw.github.io/@/flexole-dl</span> <span
                                    class="pipe">|</span>
                                <span class="env">sh</span>
                            </div>
                        </div>

                        <div class="block">
                            <div class="line comment">
                                Download byrate - a speed testing tool.
                                <i class="linux"></i><i class="macos"></i><i class="gitbash"></i>
                            </div>
                            <div class="line command"><span class="cmd">curl</span> <span class="opts">-sL</span>
                                <span class="url">https://dipakw.github.io/@/byrate-dl</span> <span
                                    class="pipe">|</span>
                                <span class="env">sh</span>
                            </div>
                        </div>

                        <div class="block">
                            <div class="line comment">
                                Download kriptun - a secure proxy server.
                                <i class="linux"></i><i class="macos"></i><i class="gitbash"></i>
                            </div>
                            <div class="line command"><span class="cmd">curl</span> <span class="opts">-sL</span>
                                <span class="url">https://dipakw.github.io/@/kriptun-dl</span> <span
                                    class="pipe">|</span>
                                <span class="env">sh</span>
                            </div>
                        </div>

                        <div class="block">
                            <div class="line comment">
                                Set up Android SDK with NDK. <i class="linux"></i>
                            </div>
                            <div class="line command"><span class="cmd">curl</span> <span class="opts">-sL</span>
                                <span class="url">https://dipakw.github.io/@/setup-android-sdk</span> <span
                                    class="pipe">|</span>
                                <span class="env">bash</span>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
    </main>

    <script type="text/javascript" src="https://cdn.jsdelivr.net/npm/toastify-js"></script>
    <script src="/script.js"></script>

</body>

</html>