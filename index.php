<?php
// Path library sesuai struktur server /public_html
require_once __DIR__ . '/../../vendor/autoload.php';

$message = '';
$extractedLinks = [];

// Ambil Kategori Langsung dari Database Maktabah
$maktabahCategories = [];
try {
    $dbConfigPath = __DIR__ . '/../maktabah.quizb.my.id/app/Config/Database.php';
    if (file_exists($dbConfigPath)) {
        require_once $dbConfigPath;
        $pdo = \App\Config\Database::getConnection();
        $stmt = $pdo->query("SELECT id, name, level FROM categories ORDER BY catord ASC, name ASC");
        $allCats = $stmt->fetchAll(PDO::FETCH_ASSOC);
        
        foreach ($allCats as $cat) {
            $prefix = str_repeat('-- ', (int)$cat['level']);
            $maktabahCategories[] = [
                'id' => $cat['id'],
                'name' => $prefix . $cat['name']
            ];
        }
    }
} catch (Exception $e) {
    // Abaikan jika error
}

// Fungsi bantuan untuk Absolute URL
function resolveUrl($baseUrl, $relativeUrl) {
    if (parse_url($relativeUrl, PHP_URL_SCHEME) != '') return $relativeUrl;
    if (strpos($relativeUrl, '//') === 0) return 'https:' . $relativeUrl;
    
    $parts = parse_url($baseUrl);
    $scheme = isset($parts['scheme']) ? $parts['scheme'] : 'http';
    $host = isset($parts['host']) ? $parts['host'] : '';
    
    if (strpos($relativeUrl, '/') === 0) {
        return $scheme . '://' . $host . $relativeUrl;
    }
    
    $path = isset($parts['path']) ? $parts['path'] : '/';
    $path = preg_replace('#/[^/]*$#', '', $path);
    return $scheme . '://' . $host . $path . '/' . $relativeUrl;
}

// Crawler Teks Rekursif untuk mempertahankan struktur teks di dalam div
function extractTextWithNewlines($node) {
    $text = '';
    
    if ($node->nodeType === XML_TEXT_NODE) {
        $text .= preg_replace('/\s+/', ' ', $node->nodeValue);
    } 
    else if ($node->nodeType === XML_ELEMENT_NODE) {
        $tagName = strtolower($node->nodeName);
        $blockTags = ['p', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'li', 'blockquote', 'tr', 'article'];
        
        if ($tagName === 'br') {
            $text .= "\n";
        }
        
        // Lewati tag h1 saat ekstraksi isi, karena h1 sudah diambil terpisah sebagai judul utama
        if (!in_array($tagName, ['script', 'style', 'nav', 'noscript', 'button', 'form', 'h1'])) {
            if ($node->hasChildNodes()) {
                foreach ($node->childNodes as $childNode) {
                    $text .= extractTextWithNewlines($childNode);
                }
            }
        }
        
        if (in_array($tagName, $blockTags)) {
            $text .= "\n";
        }
    }
    return $text;
}

// ---------------------------------------------------------
// FASE 2: PROSES DOWNLOAD DOCX
// ---------------------------------------------------------
if (!empty($_GET['download'])) {
    $url = filter_var($_GET['download'], FILTER_SANITIZE_URL);
    $html = @file_get_contents($url);
    
    if ($html) {
        $dom = new DOMDocument();
        libxml_use_internal_errors(true);
        @$dom->loadHTML(mb_convert_encoding($html, 'HTML-ENTITIES', 'UTF-8'));
        libxml_clear_errors();
        
        $xpath = new DOMXPath($dom);

        // --- FIXED: Mengambil h1 khusus yang ada di dalam content_box ---
        $h1ContentBoxNodes = $xpath->query('//div[contains(@class, "content_box")]//h1');
        $documentTitle = '';

        if ($h1ContentBoxNodes->length > 0) {
            $documentTitle = trim(strip_tags($h1ContentBoxNodes->item(0)->nodeValue));
        }

        // Fallback jika ternyata di dalam content_box tidak ada h1, cari h1 global atau h2
        if (empty($documentTitle)) {
            $fallbackNodes = $xpath->query('//h1 | //h2 | //title');
            $documentTitle = $fallbackNodes->length > 0 ? trim(strip_tags($fallbackNodes->item(0)->nodeValue)) : 'Artikel Unduhan ' . time();
        }

        // Inisiasi Word Document
        $phpWord = new \PhpOffice\PhpWord\PhpWord();
        $section = $phpWord->addSection();
        
        // Tambahkan Judul Utama di halaman Word
        $phpWord->addTitleStyle(1, ['bold' => true, 'size' => 16]);
        $section->addTitle($documentTitle, 1);
        $section->addTextBreak(1);
        
        $hasContent = false;
        $contentBoxNodes = $xpath->query('//div[contains(@class, "content_box")]');

        if ($contentBoxNodes->length > 0) {
            foreach ($contentBoxNodes as $boxNode) {
                $rawText = extractTextWithNewlines($boxNode);
                $lines = explode("\n", $rawText);

                foreach ($lines as $line) {
                    $cleanLine = trim(html_entity_decode($line, ENT_QUOTES, 'UTF-8'));
                    
                    if (!empty($cleanLine) && strlen($cleanLine) > 5) {
                        if (stripos($cleanLine, 'Share this') === false && stripos($cleanLine, 'Related Post') === false && stripos($cleanLine, 'Kirimkan Ini lewat Email') === false) {
                            $section->addText($cleanLine);
                            $section->addTextBreak(1);
                            $hasContent = true;
                        }
                    }
                }
                $section->addTextBreak(1);
            }
        } 

        // Proses Build File DOCX
        if ($hasContent) {
            // Bersihkan nama file dari karakter ilegal filesystem (\ / : * ? " < > |)
            $safeFileName = preg_replace('/[\/\\\:*?"<>|]/', '', $documentTitle);
            $safeFileName = preg_replace('/\s+/', ' ', trim($safeFileName));
            
            // Batasi nama file maksimal 120 karakter agar aman di OS Windows
            $safeFileName = substr($safeFileName, 0, 120) . '.docx';

            header("Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            header("Content-Disposition: attachment; filename=\"" . $safeFileName . "\"");
            header("Cache-Control: max-age=0");
            
            $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
            $objWriter->save('php://output');
            exit;
        } else {
            $message = "Gagal: Tidak dapat menemukan isi teks di dalam <div class=\"content_box\">.";
        }
    } else {
        $message = "Gagal mengakses URL postingan tersebut.";
    }
}

// ---------------------------------------------------------
// FASE 3: PROSES IMPORT KE MAKTABAH
// ---------------------------------------------------------
if ($_SERVER['REQUEST_METHOD'] === 'POST' && !empty($_POST['import_url'])) {
    $url = filter_var($_POST['import_url'], FILTER_SANITIZE_URL);
    $html = @file_get_contents($url);
    
    if ($html) {
        $dom = new DOMDocument();
        libxml_use_internal_errors(true);
        @$dom->loadHTML(mb_convert_encoding($html, 'HTML-ENTITIES', 'UTF-8'));
        libxml_clear_errors();
        
        $xpath = new DOMXPath($dom);

        $h1ContentBoxNodes = $xpath->query('//div[contains(@class, "content_box")]//h1');
        $documentTitle = '';

        if ($h1ContentBoxNodes->length > 0) {
            $documentTitle = trim(strip_tags($h1ContentBoxNodes->item(0)->nodeValue));
        }

        if (empty($documentTitle)) {
            $fallbackNodes = $xpath->query('//h1 | //h2 | //title');
            $documentTitle = $fallbackNodes->length > 0 ? trim(strip_tags($fallbackNodes->item(0)->nodeValue)) : 'Artikel Unduhan ' . time();
        }

        $rawText = '';
        $contentBoxNodes = $xpath->query('//div[contains(@class, "content_box")]');

        if ($contentBoxNodes->length > 0) {
            foreach ($contentBoxNodes as $boxNode) {
                $rawText .= extractTextWithNewlines($boxNode) . "\n\n";
            }
        }
        
        $lines = explode("\n", $rawText);
        $cleanLines = [];
        foreach ($lines as $line) {
            $cleanLine = trim(html_entity_decode($line, ENT_QUOTES, 'UTF-8'));
            if (!empty($cleanLine) && strlen($cleanLine) > 5) {
                if (stripos($cleanLine, 'Share this') === false && stripos($cleanLine, 'Related Post') === false && stripos($cleanLine, 'Kirimkan Ini lewat Email') === false) {
                    $cleanLines[] = $cleanLine;
                }
            }
        }
        
        $finalText = implode("\n\n", $cleanLines);
        
        if (!empty($finalText)) {
            // Pecah per halaman (3000 chars)
            $pages = [];
            $paragraphs = preg_split('/\n{2,}/', trim($finalText));
            $buf = '';
            foreach ($paragraphs as $para) {
                $para = trim($para);
                if ($para === '') continue;
                if (strlen($buf) + strlen($para) > 3000 && $buf !== '') {
                    $pages[] = trim($buf);
                    $buf = $para;
                } else {
                    $buf .= ($buf ? "\n\n" : '') . $para;
                }
            }
            if ($buf !== '') $pages[] = trim($buf);
            
            if (!empty($pages)) {
                $apiUrl = 'https://maktabah.quizb.my.id/api.php?action=admin_import_book';
                $payload = [
                    'title' => !empty($_POST['import_title']) ? trim($_POST['import_title']) : $documentTitle,
                    'author' => !empty($_POST['import_author']) ? trim($_POST['import_author']) : 'WebDownloader',
                    'category_id' => !empty($_POST['import_category_id']) ? (int)$_POST['import_category_id'] : 0,
                    'iso' => 'ar',
                    'pages' => $pages
                ];
                
                $ch = curl_init($apiUrl);
                curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
                curl_setopt($ch, CURLOPT_POST, true);
                curl_setopt($ch, CURLOPT_POSTFIELDS, json_encode($payload));
                curl_setopt($ch, CURLOPT_HTTPHEADER, [
                    'Content-Type: application/json'
                ]);
                curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, false);
                curl_setopt($ch, CURLOPT_SSL_VERIFYHOST, false);
                curl_setopt($ch, CURLOPT_TIMEOUT, 30);
                
                $response = curl_exec($ch);
                $httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
                curl_close($ch);
                
                if ($httpCode >= 200 && $httpCode < 300) {
                    $resData = json_decode($response, true);
                    if (!empty($resData['success'])) {
                        $message = "Berhasil import \"{$documentTitle}\" ke Maktabah. ID Kitab: {$resData['bkid']}, Total Halaman: {$resData['pages']}.";
                    } else {
                        $message = "Gagal import: " . ($resData['error'] ?? 'Unknown error.');
                    }
                } else {
                    $message = "Gagal terhubung ke API Maktabah. HTTP Code: {$httpCode}. Response: {$response}";
                }
            } else {
                $message = "Gagal: Teks hasil pembersihan kosong.";
            }
        } else {
            $message = "Gagal: Tidak dapat menemukan isi teks di dalam <div class=\"content_box\">.";
        }
    } else {
        $message = "Gagal mengakses URL postingan tersebut.";
    }
}

// ---------------------------------------------------------
// FASE 1: BROWSER / SCRAPER (Mengekstrak Link Direktori)
// ---------------------------------------------------------
if (!empty($_GET['url'])) {
    $scrapedUrl = filter_var($_GET['url'], FILTER_SANITIZE_URL);
    $html = @file_get_contents($scrapedUrl);
    
    if ($html === false) {
        $message = "Gagal mengambil halaman. Pastikan URL valid.";
    } else {
        $dom = new DOMDocument();
        libxml_use_internal_errors(true);
        @$dom->loadHTML(mb_convert_encoding($html, 'HTML-ENTITIES', 'UTF-8'));
        libxml_clear_errors();
        
        $xpath = new DOMXPath($dom);
        
        $mainContentLinks = $xpath->query('//div[contains(@id, "main") or contains(@class, "main")]//a[@href] | //a[@href]');
        $seenUrls = [];

        foreach ($mainContentLinks as $link) {
            $href = trim($link->getAttribute('href'));
            $text = trim(strip_tags($link->nodeValue));

            if (!empty($text) && !empty($href) && strpos($href, 'javascript:') !== 0 && strpos($href, '#') !== 0) {
                $fullUrl = resolveUrl($scrapedUrl, $href);
                $fullUrl = strtok($fullUrl, '?'); 

                if (!in_array($fullUrl, $seenUrls)) {
                    $seenUrls[] = $fullUrl;
                    $isCategory = true;
                    
                    if (preg_match('/\/\d{4}\/\d{2}\/.*\.html$/i', $fullUrl)) {
                        $isCategory = false;
                    } 
                    elseif (preg_match('/\.(jpg|jpeg|png|gif|pdf)$/i', $fullUrl)) {
                        continue;
                    }

                    $displayTitle = strlen($text) > 80 ? substr($text, 0, 80) . '...' : $text;
                    
                    $extractedLinks[] = [
                        'title' => $displayTitle,
                        'url' => $fullUrl,
                        'is_category' => $isCategory
                    ];
                }
            }
        }
        
        if (empty($extractedLinks)) {
            $message = "Tidak ada tautan yang ditemukan di halaman tersebut.";
        }
    }
}
?>

<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Web to DOCX Crawler | QuizB</title>
    <style>
        body { font-family: system-ui, -apple-system, sans-serif; background: #f3f4f6; padding: 2rem; color: #333; }
        .container { max-width: 900px; margin: 0 auto; background: white; padding: 2rem; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        .header-title { color: #2563eb; margin-bottom: 0.5rem; }
        .sub-text { color: #6b7280; font-size: 0.9rem; margin-bottom: 1.5rem; }
        .input-group { display: flex; gap: 10px; margin-bottom: 1.5rem; }
        input[type="url"] { flex: 1; padding: 0.75rem; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; }
        button { background: #2563eb; color: white; padding: 0.75rem 1.5rem; border: none; border-radius: 4px; cursor: pointer; font-weight: bold; }
        button:hover { background: #1d4ed8; }
        .alert { background: #fee2e2; color: #991b1b; padding: 1rem; border-radius: 4px; margin-bottom: 1rem; }
        .info-box { background: #dbeafe; color: #1e40af; padding: 1rem; border-radius: 4px; margin-bottom: 1rem; font-size: 0.9rem; }
        
        .link-list { list-style: none; padding: 0; margin-top: 1rem; }
        .link-item { 
            display: flex; 
            justify-content: space-between; 
            align-items: center; 
            padding: 1rem; 
            border-bottom: 1px solid #e5e7eb; 
        }
        .link-item:hover { background-color: #f9fafb; }
        .link-title { font-weight: 500; font-size: 0.95rem; display: flex; align-items: center; gap: 8px; flex: 1; padding-right: 15px;}
        .actions { display: flex; gap: 8px; flex-shrink: 0; }
        
        .badge { 
            padding: 0.4rem 0.8rem; 
            border-radius: 4px; 
            font-size: 0.8rem; 
            text-decoration: none; 
            font-weight: 600;
            text-align: center;
        }
        .badge-explore { background: #f59e0b; color: white; }
        .badge-explore:hover { background: #d97706; }
        .badge-download { background: #10b981; color: white; }
        .badge-download:hover { background: #059669; }
        .badge-import { background: #8b5cf6; color: white; }
        .badge-import:hover { background: #7c3aed; }
        .category-text { color: #f59e0b; }
        .post-text { color: #10b981; }
        
        /* Modal Styles */
        .modal-overlay { display: none; position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.5); z-index: 1000; align-items: center; justify-content: center; }
        .modal { background: white; padding: 2rem; border-radius: 8px; width: 100%; max-width: 500px; box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1); }
        .modal h3 { margin-top: 0; color: #1e3a8a; }
        .form-group { margin-bottom: 1rem; }
        .form-group label { display: block; margin-bottom: 0.5rem; font-weight: 500; font-size: 0.9rem; }
        .form-group input, .form-group select { width: 100%; padding: 0.75rem; border: 1px solid #d1d5db; border-radius: 4px; box-sizing: border-box; }
        .modal-actions { display: flex; justify-content: flex-end; gap: 10px; margin-top: 1.5rem; }
        .btn-cancel { background: #e5e7eb; color: #374151; padding: 0.75rem 1.5rem; border: none; border-radius: 4px; cursor: pointer; font-weight: bold; }
        .btn-cancel:hover { background: #d1d5db; }
        .btn-submit { background: #8b5cf6; color: white; padding: 0.75rem 1.5rem; border: none; border-radius: 4px; cursor: pointer; font-weight: bold; }
        .btn-submit:hover { background: #7c3aed; }
    </style>
</head>
<body>
    <div class="container">
        <h2 class="header-title">Web Content Box Downloader</h2>
        <p class="sub-text">Sistem ekstraksi spesifik div class="content_box" oleh <strong>webdownloader.quizb.my.id</strong></p>
        
        <?php if ($message): ?>
            <div class="alert"><?= htmlspecialchars($message) ?></div>
        <?php endif; ?>

        <form method="GET" action="">
            <div class="input-group">
                <input type="url" name="url" placeholder="Masukkan URL web atau kategori..." value="<?= htmlspecialchars($_GET['url'] ?? '') ?>" required>
                <button type="submit">Telusuri Link</button>
            </div>
        </form>

        <?php if (!empty($extractedLinks)): ?>
            <div class="info-box">
                Menemukan <strong><?= count($extractedLinks) ?></strong> tautan. <br>
                Gunakan <strong>"Buka Kategori"</strong> untuk melihat daftar isi, atau <strong>"Unduh DOCX"</strong> untuk mengekstrak struktur teks di dalamnya.
            </div>
            
            <ul class="link-list">
                <?php foreach ($extractedLinks as $item): ?>
                    <li class="link-item">
                        <span class="link-title" title="<?= htmlspecialchars($item['url']) ?>">
                            <?php if($item['is_category']): ?>
                                <span>📁</span> <span class="category-text"><?= htmlspecialchars($item['title']) ?></span>
                            <?php else: ?>
                                <span>📄</span> <span class="post-text"><?= htmlspecialchars($item['title']) ?></span>
                            <?php endif; ?>
                        </span>
                        
                        <div class="actions">
                            <?php if($item['is_category']): ?>
                                <a href="?url=<?= urlencode($item['url']) ?>" class="badge badge-explore">Buka Kategori</a>
                            <?php endif; ?>
                            
                            <?php if(!$item['is_category']): ?>
                                <?php 
                                    $domain = parse_url($item['url'], PHP_URL_HOST);
                                    if ($domain) {
                                        $domain = preg_replace('/^www\./', '', $domain);
                                    } else {
                                        $domain = 'WebDownloader';
                                    }
                                ?>
                                <button type="button" onclick="openImportModal('<?= htmlspecialchars(addslashes($item['url'])) ?>', '<?= htmlspecialchars(addslashes($item['title'])) ?>', '<?= htmlspecialchars(addslashes($domain)) ?>')" class="badge badge-import" style="border:none; cursor:pointer;">Import ke Maktabah</button>
                                <a href="?download=<?= urlencode($item['url']) ?>" class="badge badge-download" target="_blank">Unduh DOCX</a>
                            <?php endif; ?>
                        </div>
                    </li>
                <?php endforeach; ?>
            </ul>
        <?php endif; ?>
    </div>

    <!-- Modal Import -->
    <div id="importModal" class="modal-overlay">
        <div class="modal">
            <h3>Konfirmasi Import ke Maktabah</h3>
            <form method="POST" action="">
                <input type="hidden" name="import_url" id="import_url">
                
                <div class="form-group">
                    <label for="import_title">Judul Kitab / Artikel</label>
                    <input type="text" name="import_title" id="import_title" required>
                </div>
                
                <div class="form-group">
                    <label for="import_author">Penulis</label>
                    <input type="text" name="import_author" id="import_author" value="WebDownloader" required>
                </div>
                
                <div class="form-group">
                    <label for="import_category_id">Kategori</label>
                    <select name="import_category_id" id="import_category_id">
                        <option value="0">-- Pilih Kategori --</option>
                        <?php foreach ($maktabahCategories as $cat): ?>
                            <option value="<?= $cat['id'] ?>"><?= htmlspecialchars($cat['name']) ?></option>
                        <?php endforeach; ?>
                    </select>
                </div>
                
                <div class="modal-actions">
                    <button type="button" class="btn-cancel" onclick="closeImportModal()">Batal</button>
                    <button type="submit" class="btn-submit">Mulai Import</button>
                </div>
            </form>
        </div>
    </div>

    <script>
        function openImportModal(url, title, author) {
            document.getElementById('import_url').value = url;
            document.getElementById('import_title').value = title;
            document.getElementById('import_author').value = author;
            document.getElementById('importModal').style.display = 'flex';
        }
        
        function closeImportModal() {
            document.getElementById('importModal').style.display = 'none';
        }
    </script>
</body>
</html>