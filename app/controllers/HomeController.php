<?php

class HomeController {
    public function index() {
        require_once __DIR__ . '/../views/user/home.php';
    }
    public function menu() {
        require_once __DIR__ . '/../views/user/product_list.php';
    }
}

