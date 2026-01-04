<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <title>Admin Panel</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body>

<?php
$current_url = $_GET['url'] ?? '';
?>
<nav class="navbar navbar-expand-lg navbar-dark bg-dark shadow">
    <div class="container-fluid">
        <a class="navbar-brand fw-bold fs-4" href="?url=admin/users">
            ☕ Góc Cà Phê
        </a>

        <button class="navbar-toggler" type="button" data-bs-toggle="collapse" data-bs-target="#adminNavbar" aria-controls="adminNavbar" aria-expanded="false" aria-label="Toggle navigation">
            <span class="navbar-toggler-icon"></span>
        </button>

        <div class="collapse navbar-collapse" id="adminNavbar">
            <ul class="navbar-nav me-auto mb-2 mb-lg-0">
                <li class="nav-item">
                    <a class="nav-link <?= ($current_url=='admin/users')?'active':'' ?>" href="?url=admin/users">Users</a>
                </li>
                <li class="nav-item">
                    <a class="nav-link <?= ($current_url=='admin/categories') ? 'active' : '' ?>" href="?url=admin/categories">Categories</a>
                </li>
                <li class="nav-item">
                    <a class="nav-link <?= ($current_url=='admin/products')?'active':'' ?>" href="?url=admin/products">Products</a>
                </li>
                <li class="nav-item">
                    <a class="nav-link <?= ($current_url=='admin/orders')?'active':'' ?>" href="?url=admin/orders">
                        Duyệt đơn
                    </a>
                </li>
                <li class="nav-item">
                    <a class="nav-link <?= ($current_url=='admin/reservations')?'active':'' ?>" href="?url=admin/reservations">Duyệt bàn</a>
                </li>
                <li class="nav-item">
                    <a class="nav-link <?= ($current_url=='admin/revenue')?'active':'' ?>" href="?url=admin/revenue">Doanh thu</a>
                </li>

            </ul>

            <div class="d-flex align-items-center">
                <span class="text-white me-3">
                    <?= htmlspecialchars($_SESSION['user']['name'] ?? '') ?>
                </span>
                <a href="?url=logout" class="btn btn-outline-light btn-sm">Logout</a>
            </div>
        </div>
    </div>
</nav>


<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
