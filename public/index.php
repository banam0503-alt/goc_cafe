
<?php
ini_set('display_errors', 1);
ini_set('display_startup_errors', 1);
error_reporting(E_ALL);

error_reporting(E_ALL);
ini_set('display_errors', 1);
require_once __DIR__ . '/../core/bootstrap.php';
// THÊM DÒNG NÀY ĐỂ LOAD THƯ VIỆN
require_once __DIR__ . '/../vendor/autoload.php';

$url = $_GET['url'] ?? '';

switch ($url) {
    
    case '':
        (new HomeController)->index();
        break;

    case 'cart/add':
        requireLogin();
        (new CartController)->add();
        break;

    case 'cart/checkout':
        requireLogin();
        (new CartController)->checkout();
        break;

    case 'cart/payment':
        requireLogin();
        (new CartController)->payment();
        break;

    case 'cart/count':
    (new CartController)->count();
    break;
    case 'cart/delete':
        requireLogin();
        (new CartController)->delete();
        break;

    case 'cart/updateQuantity':
        requireLogin();
        (new CartController)->updateQuantity();
        break;

    case 'cart':
        requireLogin();
        (new CartController)->index();
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

    case 'datban':
        (new HomeController)->datban();
    break;

    case 'reservation/store':
        (new ReservationController)->store();
    break;
        break;
    
    case 'gioithieu':
    (new PageController)->gioithieu();
    break;


    case 'admin/staff':
        requireRole('ADMIN');
        // require_once __DIR__ . '/../controllers/AdminStaffController.php'; 
        (new AdminStaffController)->index(); // Xem danh sách & lịch
        break;

    case 'admin/staff/store':
        requireRole('ADMIN');
        (new AdminStaffController)->storeSchedule(); // Lưu lịch mới
        break;

    case 'admin/staff/delete':
        requireRole('ADMIN');
        (new AdminStaffController)->deleteSchedule(); // Xóa lịch
        break;

    /* --- KHU VỰC CỦA NHÂN VIÊN (STAFF) --- */
    case 'staff':
        // Trang dashboard chính của nhân viên
        (new StaffController)->index(); 
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
    
    /* ADMIN - RESERVATIONS */

    // 1. Form đặt bàn
    case 'reservation/create':
        require_once __DIR__ . '/../app/controllers/ReservationController.php';
        (new ReservationController)->create();
        break;

    // 2. Xử lý lưu
    case 'reservation/store':
        require_once __DIR__ . '/../app/controllers/ReservationController.php';
        (new ReservationController)->store();
        break;

    // 3. Xem lịch sử
    case 'reservation/history':
        require_once __DIR__ . '/../app/controllers/ReservationController.php';
        (new ReservationController)->history();
        break;

    case 'admin/reservations':
        requireStaffOrAdmin();
        (new AdminReservationController)->index();
    break;

    case 'admin/reservations/approve':
        requireStaffOrAdmin();
        (new AdminReservationController)->approve();
    break;

    case 'admin/reservations/cancel':
        requireStaffOrAdmin();
        (new AdminReservationController)->cancel();
    break;
        break;
    
     /* ADMIN - REVENUE MANAGEMENT */
    case 'admin/revenues':   // hỗ trợ cả 2 URL
        requireRole('ADMIN');
        (new AdminRevenueController)->index();
        break;


    case 'register':
        (new AuthController)->register();
        break;

    case 'register-handle':
        (new AuthController)->handleRegister();
        break;

    case 'profile':
        requireLogin();
        (new UserController)->profile();
        break;

    case 'profile/update':
        requireLogin();
        (new UserController)->updateProfile();
        break;

    // case 'profile/update-password':
    //     requireLogin();
    //     (new UserController)->updatePassword();
    //     break;

    /* ORDER APPROVAL - ADMIN + STAFF */

    case 'admin/orders/export-excel':
    (new OrderController())->exportExcel();
    break;



    case 'admin/orders':
        requireStaffOrAdmin();
        (new OrderController)->index();
        break;

    case 'admin/orders/approve':
        requireStaffOrAdmin();
        (new OrderController)->approve();
        break;

    case 'admin/orders/reject':
        requireStaffOrAdmin();
        (new OrderController)->reject();
        break;


    default:
        http_response_code(404);
        echo "404 - Không tìm thấy trang";

    
}

