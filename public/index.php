<?php
error_reporting(E_ALL);
ini_set('display_errors', 1);
require_once __DIR__ . '/../core/bootstrap.php';

$url = $_GET['url'] ?? '';

switch ($url) {

    case '':
        (new HomeController)->index();
        break;

    case '':
        (new HomeController)->index();
        break;
        
    case 'cart/add':
        requireLogin();
        (new CartController)->add();
        break;

    /* AUTH — KHÔNG ĐƯỢC CHẶN */
    case 'login':
        (new AuthController)->login();
        break;

    case 'login-handle':
        (new AuthController)->handleLogin();
        break;

    case 'logout':
        (new AuthController)->logout();
        break;

    /* USER */
    case 'menu':
        requireLogin();
        (new ProductController)->list();
        break;
    
    case 'gioithieu':
    (new PageController)->gioithieu();
    break;

    /* STAFF */
    case 'staff':
        requireRole('STAFF');
        echo "Trang nhân viên";
        break;

    /* ADMIN - USER MANAGEMENT */
    case 'admin':
        requireRole('ADMIN');
        (new AdminUserController)->index();
        break;

    case 'admin/users':
        requireRole('ADMIN');
        (new AdminUserController)->index();
        break;

    case 'admin/users/create':
        requireRole('ADMIN');
        (new AdminUserController)->create();
        break;

    case 'admin/users/store':
        requireRole('ADMIN');
        (new AdminUserController)->store();
        break;

    case 'admin/users/edit':
        requireRole('ADMIN');
        (new AdminUserController)->edit();
        break;

    case 'admin/users/update':
        requireRole('ADMIN');
        (new AdminUserController)->update();
        break;

    case 'admin/users/delete':
        requireRole('ADMIN');
        (new AdminUserController)->delete();
        break;

    /* ADMIN - CATEGORY MANAGEMENT */
    case 'admin/categories':
        requireRole('ADMIN');
        (new AdminCategoryController)->index();
        break;

    case 'admin/categories/create':
        requireRole('ADMIN');
        (new AdminCategoryController)->create();
        break;

    case 'admin/categories/store':
        requireRole('ADMIN');
        (new AdminCategoryController)->store();
        break;

    case 'admin/categories/edit':
        requireRole('ADMIN');
        (new AdminCategoryController)->edit();
        break;

    case 'admin/categories/update':
        requireRole('ADMIN');
        (new AdminCategoryController)->update();
        break;

    case 'admin/categories/delete':
        requireRole('ADMIN');
        (new AdminCategoryController)->delete();
        break;
    
    /* ADMIN - PRODUCT MANAGEMENT */
    case 'admin/products':
        requireRole('ADMIN');
        (new AdminProductController)->index();
        break;

    case 'admin/products/create':
        requireRole('ADMIN');
        (new AdminProductController)->create();
        break;

    case 'admin/products/store':
        requireRole('ADMIN');
        (new AdminProductController)->store();
        break;

    case 'admin/products/edit':
        requireRole('ADMIN');
        (new AdminProductController)->edit();
        break;

    case 'admin/products/update':
        requireRole('ADMIN');
        (new AdminProductController)->update();
        break;

    case 'admin/products/delete':
        requireRole('ADMIN');
        (new AdminProductController)->delete();
        break;
    
     /* ADMIN - REVENUE MANAGEMENT */
    case 'admin/revenue':
    case 'admin/revenues':   // hỗ trợ cả 2 URL
        requireRole('ADMIN');
        (new AdminRevenueController)->index();
        break;

    case 'admin/revenue/create':
        requireRole('ADMIN');
        (new AdminRevenueController)->create();
        break;

    case 'admin/revenue/store':
        requireRole('ADMIN');
        (new AdminRevenueController)->store();
        break;

    case 'admin/revenue/edit':
        requireRole('ADMIN');
        (new AdminRevenueController)->edit();
        break;

    case 'admin/revenue/update':
        requireRole('ADMIN');
        (new AdminRevenueController)->update();
        break;

    case 'admin/revenue/delete':
        requireRole('ADMIN');
        (new AdminRevenueController)->delete();
        break;

case 'register':
    (new AuthController)->register();
    break;

case 'register-handle':
    (new AuthController)->handleRegister();
    break;

    default:
        http_response_code(404);
        echo "404 - Không tìm thấy trang";

    
}
