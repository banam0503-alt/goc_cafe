<?php
class CartController
{
    public static function getCartCount($userId)
    {
        $pdo = Database::connect();

        $stmt = $pdo->prepare("
            SELECT SUM(oi.quantity)
            FROM orders o
            JOIN order_items oi ON o.id = oi.order_id
            WHERE o.user_id = ? AND o.status = 'PENDING'
        ");
        $stmt->execute([$userId]);

        return $stmt->fetchColumn() ?? 0;
    }

    public function add() {
        $pdo = Database::connect();
        $userId = $_SESSION['user']['id'];
        $productId = $_POST['product_id'];

        // tìm order PENDING
        $stmt = $pdo->prepare("SELECT id FROM orders WHERE user_id=? AND status='PENDING' LIMIT 1");
        $stmt->execute([$userId]);
        $order = $stmt->fetch();

        $orderId = $order ? $order['id'] : null;
        if (!$orderId) {
            $stmt = $pdo->prepare("INSERT INTO orders (user_id, total_price) VALUES (?, 0)");
            $stmt->execute([$userId]);
            $orderId = $pdo->lastInsertId();
        }

        // kiểm tra sản phẩm
        $stmt = $pdo->prepare("SELECT id, quantity FROM order_items WHERE order_id=? AND product_id=?");
        $stmt->execute([$orderId, $productId]);
        $item = $stmt->fetch();

        if($item){
            $stmt = $pdo->prepare("UPDATE order_items SET quantity = quantity + 1 WHERE id=?");
            $stmt->execute([$item['id']]);
        } else {
            $stmt = $pdo->prepare("SELECT price FROM products WHERE id=?");
            $stmt->execute([$productId]);
            $price = $stmt->fetchColumn();

            $stmt = $pdo->prepare("INSERT INTO order_items (order_id, product_id, quantity, price) VALUES (?, ?, 1, ?)");
            $stmt->execute([$orderId, $productId, $price]);
        }

        // trả số lượng giỏ hàng
        $stmt = $pdo->prepare("SELECT SUM(oi.quantity) FROM orders o JOIN order_items oi ON o.id=oi.order_id WHERE o.user_id=? AND o.status='PENDING'");
        $stmt->execute([$userId]);
        $cartCount = $stmt->fetchColumn() ?? 0;

        header('Content-Type: application/json');
        echo json_encode(['success'=>true, 'cartCount'=>$cartCount]);
        exit;
    }


public function index()
{
    $pdo = Database::connect();
    $userId = $_SESSION['user']['id'];

    // lọc theo trạng thái order
    $status = $_GET['status'] ?? 'PENDING';
    if (!in_array($status, ['PENDING', 'PAID', 'APPROVED', 'CANCELLED'])) {
        $status = 'PENDING';
    }

    $stmt = $pdo->prepare("
        SELECT 
            oi.id AS order_item_id,
            p.id AS product_id,
            p.name,
            p.image,
            oi.quantity,
            oi.price,

            o.status AS order_status
        FROM order_items oi
        JOIN orders o ON o.id = oi.order_id
        JOIN products p ON p.id = oi.product_id
        WHERE o.user_id = ?
          AND o.status = ?
        ORDER BY oi.id DESC
    ");

    $stmt->execute([$userId, $status]);
    $items = $stmt->fetchAll(PDO::FETCH_ASSOC);

    require_once __DIR__ . '/../views/user/cart.php';
}




// Trang nhập địa chỉ + chọn phương thức thanh toán
// 1. Hiển thị trang Checkout
public function checkout()
{
    // Kiểm tra có chọn món chưa
    $ids = $_GET['items'] ?? '';
    if (!$ids) {
        header('Location: ?url=cart');
        exit;
    }

    $pdo = Database::connect();
    $userId = $_SESSION['user']['id'];

    // [MỚI] Lấy thông tin User hiện tại để check xem có SĐT/Địa chỉ chưa
    $userStmt = $pdo->prepare("SELECT name, phone, address FROM users WHERE id = ?");
    $userStmt->execute([$userId]);
    $currentUser = $userStmt->fetch(PDO::FETCH_ASSOC);

    // Lấy danh sách món hàng
    $ids = array_map('intval', explode(',', $ids));
    // Dùng implode an toàn hơn cho SQL IN
    $inQuery = implode(',', array_fill(0, count($ids), '?'));

    $stmt = $pdo->prepare("
        SELECT 
            oi.id,
            p.name,
            oi.quantity,
            oi.price
        FROM order_items oi
        JOIN orders o ON o.id = oi.order_id
        JOIN products p ON p.id = oi.product_id
        WHERE oi.id IN ($inQuery)
          AND o.status = 'PENDING'
    ");
    
    // Execute với mảng ids
    $stmt->execute($ids);
    $items = $stmt->fetchAll(PDO::FETCH_ASSOC);

    require __DIR__ . '/../views/user/checkout.php';
}


public function delete()
{
    $id = $_POST['id'] ?? 0;
    $userId = $_SESSION['user']['id'];

    if (!$id) {
        echo json_encode(['success'=>false]);
        exit;
    }

    $pdo = Database::connect();

    $stmt = $pdo->prepare("
        DELETE oi
        FROM order_items oi
        JOIN orders o ON o.id = oi.order_id
        WHERE oi.id = ?
          AND o.user_id = ?
          AND o.status = 'PENDING'
    ");

    $stmt->execute([$id, $userId]);

    echo json_encode([
    'success' => $stmt->rowCount() > 0
]);
    exit;
}




// 2. Xử lý Thanh toán
public function payment()
{
    $userId = $_SESSION['user']['id'];
    $items = array_map('intval', explode(',', $_POST['items'] ?? ''));
    
    // Lấy thông tin từ Form
    $address = trim($_POST['address'] ?? '');
    $phone = trim($_POST['phone'] ?? '');
    $method = $_POST['payment_method'] ?? 'momo';

    if (empty($items) || empty($address) || empty($phone)) {
        echo "<script>alert('Vui lòng nhập đầy đủ địa chỉ và số điện thoại!'); window.history.back();</script>";
        exit;
    }
    
    $pdo = Database::connect();
    $pdo->beginTransaction();

    try {
        // [QUAN TRỌNG] Cập nhật lại thông tin User nếu họ nhập mới/thay đổi lúc thanh toán
        // Vì bảng orders không có cột address/phone, ta lưu vào user để dùng cho lần sau
        $updUser = $pdo->prepare("UPDATE users SET phone = ?, address = ? WHERE id = ?");
        $updUser->execute([$phone, $address, $userId]);

        // Cập nhật lại session để hiển thị đúng ngay lập tức (nếu cần dùng ở chỗ khác)
        $_SESSION['user']['phone'] = $phone;
        $_SESSION['user']['address'] = $address;

        // 1️⃣ Tạo order mới với trạng thái PAID (hoặc PENDING chờ duyệt tùy logic)
        $stmt = $pdo->prepare("
            INSERT INTO orders (user_id, status, total_price, created_at)
            VALUES (?, 'PAID', 0, NOW())
        ");
        $stmt->execute([$userId]);
        $newOrderId = $pdo->lastInsertId();

        // 2️⃣ Chuyển các items đã chọn sang order mới này
        // Cần tạo placeholder cho câu lệnh IN
        $inQuery = implode(',', array_fill(0, count($items), '?'));
        
        // Merge tham số: order_id mới + danh sách item ids
        $params = array_merge([$newOrderId], $items);

        $stmt = $pdo->prepare("
            UPDATE order_items
            SET order_id = ?
            WHERE id IN ($inQuery)
        ");
        $stmt->execute($params);

        // 3️⃣ Tính tổng tiền của Order mới
        $stmt = $pdo->prepare("
            SELECT SUM(quantity * price)
            FROM order_items
            WHERE order_id = ?
        ");
        $stmt->execute([$newOrderId]);
        $totalPrice = $stmt->fetchColumn() ?? 0;

        // 4️⃣ Update total_price cho Order
        $stmt = $pdo->prepare("
            UPDATE orders
            SET total_price = ?
            WHERE id = ?
        ");
        $stmt->execute([$totalPrice, $newOrderId]);

        // 5️⃣ Lưu session để hiển thị trang Success
        $_SESSION['last_order'] = [
            'id' => $newOrderId,
            'address' => $address,
            'phone' => $phone,
            'method' => $method,
            'total' => $totalPrice
        ];

        $pdo->commit();

        require __DIR__ . '/../views/user/payment_success.php';

    } catch (Exception $e) {
        $pdo->rollBack();
        // Ghi log lỗi thực tế để debug: error_log($e->getMessage());
        echo "Lỗi thanh toán: " . $e->getMessage();
    }
}



public function count()
{
    if (!isset($_SESSION['user'])) {
        echo json_encode(['count' => 0]);
        exit;
    }

    $userId = $_SESSION['user']['id'];
    $count = self::getCartCount($userId);

    echo json_encode(['count' => $count]);
    exit;
}

public function updateQuantity()
{
    $id = $_POST['id'] ?? 0;
    $type = $_POST['type'] ?? '';
    $userId = $_SESSION['user']['id'];

    if (!$id || !in_array($type, ['inc','dec'])) {
        echo json_encode(['success'=>false]);
        exit;
    }

    $pdo = Database::connect();

    if ($type === 'inc') {
        $stmt = $pdo->prepare("
            UPDATE order_items oi
            JOIN orders o ON o.id = oi.order_id
            SET oi.quantity = oi.quantity + 1
            WHERE oi.id=? AND o.status='PENDING' AND o.user_id=?
        ");
    } else {
        $stmt = $pdo->prepare("
            UPDATE order_items oi
            JOIN orders o ON o.id = oi.order_id
            SET oi.quantity = GREATEST(1, oi.quantity - 1)
            WHERE oi.id=? AND o.status='PENDING' AND o.user_id=?
        ");
    }

    $stmt->execute([$id, $userId]);

    $stmt = $pdo->prepare("SELECT quantity FROM order_items WHERE id=?");
    $stmt->execute([$id]);
    $qty = $stmt->fetchColumn();

    echo json_encode(['success'=>true, 'quantity'=>$qty]);
    exit;
}


}


