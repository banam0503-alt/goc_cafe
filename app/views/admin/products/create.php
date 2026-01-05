<?php require_once __DIR__ . '/../../layouts/admin_header.php'; ?>

<div class="container mt-4">
    <h2>Thêm Product</h2>
    <form method="POST" action="?url=admin/products/store" enctype="multipart/form-data">
        <div class="mb-3">
            <label class="form-label">Tên Product</label>
            <input type="text" name="name" class="form-control" required>
        </div>
        <div class="mb-3">
            <label class="form-label">Category</label>
            <select name="category_id" class="form-select" required>
                <option value="">-- Chọn Category --</option>
                <?php foreach ($categories as $cat): ?>
                    <option value="<?= $cat['id'] ?>"><?= htmlspecialchars($cat['name']) ?></option>
                <?php endforeach ?>
            </select>
        </div>
        <div class="mb-3">
            <label class="form-label">Price</label>
            <input type="number" name="price" class="form-control" required>
        </div>
        <div class="mb-3">
            <label class="form-label">Status</label>
            <select name="status" class="form-select">
                <option value="AVAILABLE">AVAILABLE</option>
                <option value="HIDDEN">HIDDEN</option>
                <option value="SPECIAL">SPECIAL</option>
            </select>
        </div>
        <div class="mb-3">
            <label class="form-label">Image</label>
            <input type="file" name="image" class="form-control">
        </div>
        <div class="mb-3">
            <label class="form-label">Description</label>
            <textarea name="description" class="form-control"></textarea>
        </div>
        <button type="submit" class="btn btn-success">➕ Thêm</button>
        <a href="?url=admin/products" class="btn btn-secondary">Hủy</a>
    </form>
</div>
