<?php
require_once __DIR__ . '/../layouts/header.php';
?>

<!-- GIỚI THIỆU QUÁN -->
<section class="intro-section">
    <div class="intro-wrapper">
        <div class="intro-text">
            <h2>Giới thiệu quán Góc Cafe</h2>

            <p>
                Góc Cafe được thành lập với mong muốn mang lại một không gian thư giãn,
                nhẹ nhàng và gần gũi cho mọi khách hàng. Chúng tôi luôn chú trọng vào
                chất lượng từng ly cà phê rang xay nguyên chất.
            </p>

            <p>
                Với phương châm <b>“Chất lượng tạo nên trải nghiệm”</b>, Góc Cafe cam kết
                mang đến hương vị cà phê mộc mạc, thơm nồng cùng sự phục vụ chu đáo.
            </p>

            <p>
                Ngoài cà phê, quán còn phục vụ trà trái cây, nước ép tươi và bánh ngọt homemade.
            </p>
        </div>

        <div class="intro-image">
            <img src="/GocCaPhe/public/assets/img/bgg-coffee.jpg" alt="">
        </div>
    </div>
</section>

<!-- DỊCH VỤ -->
<section class="service-section">
    <h2>Dịch vụ của Góc Cafe</h2>
    <div class="service-container">

        <div class="service-box">
            <img src="/GocCaPhe/public/assets/img/icon-coffee.png">
            <h3>Cà phê rang xay nguyên chất</h3>
            <p>Rang mới mỗi ngày, giữ trọn hương vị.</p>
        </div>

        <div class="service-box">
            <img src="/GocCaPhe/public/assets/img/icon-machine.png">
            <h3>Máy pha hiện đại</h3>
            <p>Pha chế chuẩn vị quốc tế.</p>
        </div>

        <div class="service-box">
            <img src="/GocCaPhe/public/assets/img/icon-dessert.png">
            <h3>Đồ ngọt & trà trái cây</h3>
            <p>Bánh homemade và nước trái cây tươi.</p>
        </div>

        <div class="service-box">
            <img src="/GocCaPhe/public/assets/img/icon-delivery.png">
            <h3>Giao hàng tận nơi</h3>
            <p>Hỗ trợ giao hàng nội thành.</p>
        </div>

    </div>
</section>

<!-- THÀNH TỰU -->
<section class="achievement-section">
    <h2>Thành tựu</h2>
    <div class="achievement-container">
        <div class="achievement-box"><h3>2023</h3><p>Top 10 quán cafe phong cách.</p></div>
        <div class="achievement-box"><h3>2024</h3><p>100.000+ khách hàng/năm.</p></div>
        <div class="achievement-box"><h3>2021</h3><p>Dòng cà phê rang mộc độc quyền.</p></div>
        <div class="achievement-box"><h3>2022</h3><p>Chứng nhận VSATTP 5 sao.</p></div>
    </div>
</section>

<!-- VIDEO -->
<section class="coffee-video">
    <h2>Quy trình pha cà phê</h2>
    <video controls>
        <source src="/GocCaPhe/video.mp4" type="video/mp4">
    </video>
</section>

<!-- ROBUSTA -->
<section class="coffee-robusta">
    <h2>Hạt cà phê Robusta</h2>
    <p>
        Robusta có hàm lượng caffeine cao, vị đậm, ít chua.
        Việt Nam là một trong những nước xuất khẩu Robusta lớn nhất thế giới.
    </p>
    <img src="/GocCaPhe/public/assets/img/robusta.jpg">
</section>
<?php require_once __DIR__ . '/../layouts/footer.php'; ?>
