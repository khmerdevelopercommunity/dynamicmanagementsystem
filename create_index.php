<?php
// Ensure uploads folder exists upfront
$uploads_dir = __DIR__ . '/uploads';
if (!is_dir($uploads_dir)) {
    mkdir($uploads_dir, 0777, true);
}

// Handle Builder Setup Submission
if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_POST['action_builder'])) {
    $fields = $_POST['fields'] ?? [];
    
    // Filter empty values and remove duplicate dynamic column names
    $clean_fields = [];
    foreach ($fields as $f) {
        $trimmed = trim($f);
        if (!empty($trimmed) && !in_array($trimmed, $clean_fields, true)) {
            $clean_fields[] = $trimmed;
        }
    }

    $system_name = !empty($_POST['system_name']) ? trim($_POST['system_name']) : 'Management System';

    // Process optional header logo (File OR URL)
    $system_logo = '';
    if (isset($_FILES['system_logo']) && $_FILES['system_logo']['error'] === UPLOAD_ERR_OK) {
        $ext = strtolower(pathinfo($_FILES['system_logo']['name'], PATHINFO_EXTENSION));
        $allowed = ['jpg', 'jpeg', 'png', 'gif', 'webp'];
        if (in_array($ext, $allowed)) {
            $logo_filename = 'logo_' . uniqid() . '.' . $ext;
            if (move_uploaded_file($_FILES['system_logo']['tmp_name'], $uploads_dir . '/' . $logo_filename)) {
                $system_logo = 'uploads/' . $logo_filename;
            }
        }
    } elseif (!empty($_POST['system_logo_url'])) {
        $system_logo = trim($_POST['system_logo_url']);
    }

    // 1. CREATE DATABASE.DB (SQLite3)
    $db_file = __DIR__ . '/database.db';
    if (file_exists($db_file)) {
        unlink($db_file);
    }
    
    $db = new SQLite3($db_file);
    $db->exec('PRAGMA encoding = "UTF-8";');
    $db->exec('PRAGMA foreign_keys = ON;');

    $columns = [
        'id INTEGER PRIMARY KEY AUTOINCREMENT',
        'avatar TEXT',
        'created_at DATETIME DEFAULT CURRENT_TIMESTAMP'
    ];
    
    foreach ($clean_fields as $field) {
        $columns[] = '"' . str_replace('"', '""', $field) . '" TEXT';
    }

    $query = 'CREATE TABLE IF NOT EXISTS records (' . implode(', ', $columns) . ');';
    $db->exec($query);
    $db->close();

    // 2. GENERATE INDEX.PHP
    $schema_export = var_export($clean_fields, true);
    $logo_export = var_export($system_logo, true);
    $name_export = var_export($system_name, true);

    $index_code = '<?php
header("Content-Type: text/html; charset=utf-8");
$db = new SQLite3(__DIR__ . "/database.db");
$db->exec("PRAGMA encoding = \"UTF-8\";");

$SYSTEM_NAME = ' . $name_export . ';
$SCHEMA_FIELDS = ' . $schema_export . ';
$SYSTEM_LOGO = ' . $logo_export . ';
$error = "";

// Ensure records table exists safely
$cols = [
    \'id INTEGER PRIMARY KEY AUTOINCREMENT\',
    \'avatar TEXT\',
    \'created_at DATETIME DEFAULT CURRENT_TIMESTAMP\'
];
foreach ($SCHEMA_FIELDS as $field) {
    $cols[] = \'"\' . str_replace(\'"\', \'""\', $field) . \'" TEXT\';
}
$db->exec(\'CREATE TABLE IF NOT EXISTS records (\' . implode(\', \', $cols) . \');\');

// Dynamically check and add any missing columns without crashing on duplicates
$table_info = $db->query("PRAGMA table_info(records)");
$existing_cols = [];
if ($table_info) {
    while ($col_row = $table_info->fetchArray(SQLITE3_ASSOC)) {
        $existing_cols[] = $col_row[\'name\'];
    }
}

foreach ($SCHEMA_FIELDS as $field) {
    if (!in_array($field, $existing_cols, true)) {
        $safe_field = str_replace(\'"\', \'""\', $field);
        @$db->exec(\'ALTER TABLE records ADD COLUMN "\' . $safe_field . \'" TEXT;\');
    }
}

// Helper to remove local server files safely
function removeImageFile($path) {
    if (!empty($path) && strpos($path, "uploads/") === 0) {
        $full_path = __DIR__ . "/" . $path;
        if (file_exists($full_path) && is_file($full_path)) {
            unlink($full_path);
        }
    }
}

// ---------------- Handle EXCEL (.xls) EXPORT ----------------
if (isset($_GET["action"]) && $_GET["action"] === "export") {
    $clean_sys_name = preg_replace("/[^a-zA-Z0-9_\-]/", "_", $SYSTEM_NAME);
    $clean_sys_name = trim(preg_replace("/_+/", "_", $clean_sys_name), "_");
    if (empty($clean_sys_name)) {
        $clean_sys_name = "System";
    }

    $filename = $clean_sys_name . "_export_" . date("Y-m-d_H-i-s") . ".xls";
    
    header("Content-Type: application/vnd.ms-excel; charset=utf-8");
    header("Content-Disposition: attachment; filename=" . $filename);
    header("Pragma: no-cache");
    header("Expires: 0");

    echo \'<?xml version="1.0" encoding="UTF-8"?>\' . "\n";
    echo \'<?mso-application progid="Excel.Sheet"?>\' . "\n";
    ?>
    <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
              xmlns:o="urn:schemas-microsoft-com:office:office"
              xmlns:x="urn:schemas-microsoft-com:office:excel"
              xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
      <Styles>
        <Style ss:ID="Header">
          <Font ss:Bold="1" ss:Color="#FFFFFF"/>
          <Interior ss:Color="#0F172A" ss:Pattern="Solid"/>
          <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
        </Style>
        <Style ss:ID="Data">
          <Alignment ss:Vertical="Center"/>
        </Style>
      </Styles>
      <Worksheet ss:Name="Records">
        <Table>
          <Row ss:StyleID="Header">
            <Cell><Data ss:Type="String">ID</Data></Cell>
            <?php foreach ($SCHEMA_FIELDS as $field): ?>
              <Cell><Data ss:Type="String"><?= htmlspecialchars($field, ENT_QUOTES, "UTF-8") ?></Data></Cell>
            <?php endforeach; ?>
            <Cell><Data ss:Type="String">Profile Image Link / Path</Data></Cell>
            <Cell><Data ss:Type="String">Created At</Data></Cell>
          </Row>
          <?php
          $results = $db->query("SELECT * FROM records ORDER BY id ASC");
          if ($results):
            while ($row = $results->fetchArray(SQLITE3_ASSOC)):
                $img_display = $row["avatar"] ?? "";
                $is_url = (strpos($img_display, "http://") === 0 || strpos($img_display, "https://") === 0);
            ?>
              <Row ss:StyleID="Data">
                <Cell><Data ss:Type="Number"><?= $row["id"] ?></Data></Cell>
                <?php foreach ($SCHEMA_FIELDS as $field): ?>
                  <Cell><Data ss:Type="String"><?= htmlspecialchars($row[$field] ?? "", ENT_QUOTES, "UTF-8") ?></Data></Cell>
                <?php endforeach; ?>
                
                <?php if ($is_url): ?>
                  <Cell ss:HRef="<?= htmlspecialchars($img_display, ENT_QUOTES, "UTF-8") ?>">
                      <Data ss:Type="String"><?= htmlspecialchars($img_display, ENT_QUOTES, "UTF-8") ?></Data>
                  </Cell>
                <?php else: ?>
                  <Cell><Data ss:Type="String"><?= htmlspecialchars($img_display, ENT_QUOTES, "UTF-8") ?></Data></Cell>
                <?php endif; ?>

                <Cell><Data ss:Type="String"><?= htmlspecialchars($row["created_at"] ?? "", ENT_QUOTES, "UTF-8") ?></Data></Cell>
              </Row>
            <?php endwhile; 
          endif; ?>
        </Table>
      </Worksheet>
    </Workbook>
    <?php
    exit;
}

// ---------------- Handle DELETE ----------------
if (isset($_GET["action"]) && $_GET["action"] === "delete" && isset($_GET["id"])) {
    $id = (int)$_GET["id"];
    
    $stmt = $db->prepare("SELECT avatar FROM records WHERE id = :id");
    if ($stmt) {
        $stmt->bindValue(":id", $id, SQLITE3_INTEGER);
        $res = $stmt->execute()->fetchArray(SQLITE3_ASSOC);
        if ($res) {
            removeImageFile($res["avatar"]);
            
            $del_stmt = $db->prepare("DELETE FROM records WHERE id = :id");
            $del_stmt->bindValue(":id", $id, SQLITE3_INTEGER);
            $del_stmt->execute();
        }
    }
    header("Location: index.php");
    exit;
}

// ---------------- Handle CREATE & UPDATE ----------------
if ($_SERVER["REQUEST_METHOD"] === "POST" && isset($_POST["action"])) {
    $action = $_POST["action"];
    $record_id = isset($_POST["id"]) ? (int)$_POST["id"] : 0;
    $avatar_path = "";
    $allowed = ["jpg", "jpeg", "png", "gif", "webp"];

    $existing = null;
    if ($action === "update" && $record_id > 0) {
        $stmt = $db->prepare("SELECT * FROM records WHERE id = :id");
        if ($stmt) {
            $stmt->bindValue(":id", $record_id, SQLITE3_INTEGER);
            $existing = $stmt->execute()->fetchArray(SQLITE3_ASSOC);
            $avatar_path = $existing["avatar"] ?? "";
        }
    }

    $target_dir = __DIR__ . "/uploads/";
    if (!is_dir($target_dir)) {
        mkdir($target_dir, 0777, true);
    }

    if (isset($_FILES["avatar"]) && $_FILES["avatar"]["error"] === UPLOAD_ERR_OK) {
        $tmp_name = $_FILES["avatar"]["tmp_name"];
        $ext = strtolower(pathinfo($_FILES["avatar"]["name"], PATHINFO_EXTENSION));
        
        if (in_array($ext, $allowed)) {
            $filename = uniqid("img_", true) . "." . $ext;
            if (move_uploaded_file($tmp_name, $target_dir . $filename)) {
                if ($action === "update" && !empty($existing["avatar"])) {
                    removeImageFile($existing["avatar"]);
                }
                $avatar_path = "uploads/" . $filename;
            } else {
                $error = "Failed to upload file to local server.";
            }
        } else {
            $error = "Invalid file format! Allowed: JPG, JPEG, PNG, GIF, WEBP.";
        }
    } 
    elseif (!empty($_POST["avatar_url"])) {
        $url = trim($_POST["avatar_url"]);
        if (filter_var($url, FILTER_VALIDATE_URL)) {
            if ($action === "update" && !empty($existing["avatar"])) {
                removeImageFile($existing["avatar"]);
            }
            $avatar_path = $url;
        } else {
            $error = "Invalid Image URL provided.";
        }
    } 
    elseif ($action === "create" && empty($avatar_path)) {
        $error = "Please upload an image file or provide an Image URL.";
    }

    if (empty($error)) {
        if ($action === "create") {
            $cols = ["avatar"];
            $vals = [":avatar"];
            $params = [":avatar" => $avatar_path];

            foreach ($SCHEMA_FIELDS as $field) {
                $cols[] = \'"\' . str_replace(\'"\', \'""\', $field) . \'"\';
                $param_key = ":" . md5($field);
                $vals[] = $param_key;
                $params[$param_key] = $_POST[$field] ?? "";
            }

            $stmt = $db->prepare("INSERT INTO records (" . implode(", ", $cols) . ") VALUES (" . implode(", ", $vals) . ")");
            if ($stmt) {
                foreach ($params as $key => $val) {
                    $stmt->bindValue($key, $val, SQLITE3_TEXT);
                }
                $stmt->execute();
            }
        } 
        elseif ($action === "update" && $record_id > 0) {
            $set_clauses = ["avatar = :avatar"];
            $params = [":avatar" => $avatar_path, ":id" => $record_id];

            foreach ($SCHEMA_FIELDS as $field) {
                $param_key = ":" . md5($field);
                $set_clauses[] = \'"\' . str_replace(\'"\', \'""\', $field) . \'" = \' . $param_key;
                $params[$param_key] = $_POST[$field] ?? "";
            }

            $stmt = $db->prepare("UPDATE records SET " . implode(", ", $set_clauses) . " WHERE id = :id");
            if ($stmt) {
                foreach ($params as $key => $val) {
                    $type = ($key === ":id") ? SQLITE3_INTEGER : SQLITE3_TEXT;
                    $stmt->bindValue($key, $val, $type);
                }
                $stmt->execute();
            }
        }

        header("Location: index.php");
        exit;
    }
}

// Fetch record to edit
$edit_data = null;
if (isset($_GET["action"]) && $_GET["action"] === "edit" && isset($_GET["id"])) {
    $edit_id = (int)$_GET["id"];
    $stmt = $db->prepare("SELECT * FROM records WHERE id = :id");
    if ($stmt) {
        $stmt->bindValue(":id", $edit_id, SQLITE3_INTEGER);
        $edit_data = $stmt->execute()->fetchArray(SQLITE3_ASSOC);
    }
}

$results = $db->query("SELECT * FROM records ORDER BY id ASC");
?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title><?= htmlspecialchars($SYSTEM_NAME, ENT_QUOTES, "UTF-8") ?></title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&amp;display=swap" rel="stylesheet">
    <style>
        :root {
            --bg: #f8fafc;
            --surface: #ffffff;
            --border: #e2e8f0;
            --text-main: #0f172a;
            --text-muted: #64748b;
            --primary: #2563eb;
            --primary-hover: #1d4ed8;
            --danger: #ef4444;
            --danger-bg: #fef2f2;
            --radius: 10px;
        }

        * { box-sizing: border-box; }
        body { font-family: \'Inter\', -apple-system, BlinkMacSystemFont, sans-serif; background: var(--bg); color: var(--text-main); margin: 0; padding: 24px; line-height: 1.5; }

        .navbar { display: flex; align-items: center; justify-content: space-between; background: var(--surface); padding: 16px 24px; border-radius: var(--radius); border: 1px solid var(--border); margin-bottom: 24px; }
        .navbar-brand { display: flex; align-items: center; gap: 14px; }
        .navbar-brand img { height: 38px; border-radius: 6px; object-fit: contain; }
        .navbar-title { font-size: 18px; font-weight: 700; color: var(--text-main); margin: 0; }

        .main-grid { display: grid; grid-template-columns: 340px minmax(0, 1fr); gap: 24px; align-items: start; }
        .card { background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius); padding: 20px; overflow: hidden; }

        .card-header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 20px; }
        .card-title { font-size: 15px; font-weight: 600; color: var(--text-main); margin: 0; }

        .form-group { margin-bottom: 16px; }
        .form-label { display: block; font-size: 13px; font-weight: 500; color: var(--text-muted); margin-bottom: 6px; }
        .form-control { width: 100%; padding: 9px 12px; font-size: 14px; font-family: inherit; background: var(--surface); border: 1px solid var(--border); border-radius: 6px; color: var(--text-main); transition: border-color 0.15s ease; }
        .form-control:focus { outline: none; border-color: var(--primary); box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.1); }

        .avatar-section { display: flex; align-items: center; gap: 16px; margin-bottom: 16px; }
        .avatar-preview { width: 64px; height: 64px; border-radius: 50%; object-fit: cover; border: 1px solid var(--border); background: #f1f5f9; flex-shrink: 0; }
        .avatar-inputs { display: flex; flex-direction: column; gap: 8px; width: 100%; }

        .btn { display: inline-flex; align-items: center; justify-content: center; padding: 9px 16px; font-size: 14px; font-weight: 500; font-family: inherit; border-radius: 6px; border: 1px solid transparent; cursor: pointer; text-decoration: none; transition: all 0.15s ease; }
        .btn-primary { background: var(--primary); color: #ffffff; width: 100%; }
        .btn-primary:hover { background: var(--primary-hover); }
        .btn-secondary { background: var(--surface); border-color: var(--border); color: var(--text-main); width: 100%; margin-top: 8px; }
        .btn-secondary:hover { background: #f8fafc; }
        .btn-excel { background: #059669; color: white; padding: 7px 12px; font-size: 13px; }
        .btn-excel:hover { background: #047857; }

        /* HORIZONTAL SCROLL ENHANCEMENTS */
        .table-container { 
            width: 100%; 
            overflow-x: auto; 
            border: 1px solid var(--border); 
            border-radius: 8px; 
            -webkit-overflow-scrolling: touch;
        }
        table { width: 100%; border-collapse: collapse; min-width: max-content; font-size: 14px; }
        th { background: #f8fafc; color: var(--text-muted); font-weight: 600; font-size: 12px; text-transform: uppercase; letter-spacing: 0.05em; text-align: left; padding: 12px 16px; border-bottom: 1px solid var(--border); white-space: nowrap; }
        td { padding: 12px 16px; border-bottom: 1px solid var(--border); color: var(--text-main); vertical-align: middle; white-space: nowrap; }
        tr:last-child td { border-bottom: none; }
        tr:hover td { background: #f8fafc; }

        .avatar-thumb { width: 36px; height: 36px; border-radius: 50%; object-fit: cover; border: 1px solid var(--border); }
        .action-links { display: flex; gap: 8px; }
        .btn-action { padding: 4px 10px; font-size: 12px; font-weight: 500; border-radius: 4px; text-decoration: none; border: 1px solid var(--border); color: var(--text-main); background: var(--surface); }
        .btn-action:hover { background: #f1f5f9; }
        .btn-action-delete { color: var(--danger); border-color: #fca5a5; background: var(--danger-bg); }
        .btn-action-delete:hover { background: #fee2e2; }

        .alert { background: #fef2f2; border: 1px solid #fca5a5; color: var(--danger); padding: 12px 16px; border-radius: 6px; font-size: 14px; margin-bottom: 20px; }

        @media (max-width: 900px) {
            .main-grid { grid-template-columns: 1fr; }
        }
    </style>
</head>
<body>

<div class="navbar">
    <div class="navbar-brand">
        <?php if (!empty($SYSTEM_LOGO)): ?>
            <?php $logo_src = (strpos($SYSTEM_LOGO, "http") === 0) ? $SYSTEM_LOGO : htmlspecialchars($SYSTEM_LOGO, ENT_QUOTES, "UTF-8"); ?>
            <img src="<?= $logo_src ?>" alt="Logo">
        <?php endif; ?>
        <h1 class="navbar-title"><?= htmlspecialchars($SYSTEM_NAME, ENT_QUOTES, "UTF-8") ?></h1>
    </div>
</div>

<?php if (!empty($error)): ?>
    <div class="alert"><?= htmlspecialchars($error, ENT_QUOTES, "UTF-8") ?></div>
<?php endif; ?>

<div class="main-grid">
    <div class="card">
        <div class="card-header">
            <h2 class="card-title"><?= $edit_data ? "Edit Record" : "New Entry" ?></h2>
        </div>
        <form method="POST" enctype="multipart/form-data">
            <input type="hidden" name="action" value="<?= $edit_data ? "update" : "create" ?>">
            <?php if ($edit_data): ?>
                <input type="hidden" name="id" value="<?= $edit_data["id"] ?>">
            <?php endif; ?>

            <div class="form-group">
                <label class="form-label">Profile Image</label>
                <div class="avatar-section">
                    <?php 
                    $initial_img = "data:image/svg+xml;utf8,<svg xmlns=\'http://www.w3.org/2000/svg\' width=\'64\' height=\'64\' viewBox=\'0 0 24 24\' fill=\'%23cbd5e1\'><path d=\'M12 12c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4zm0 2c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4z\'/></svg>";
                    if ($edit_data && !empty($edit_data["avatar"])) {
                        $initial_img = htmlspecialchars($edit_data["avatar"], ENT_QUOTES, "UTF-8");
                    }
                    ?>
                    <img id="avatar-preview" src="<?= $initial_img ?>" class="avatar-preview" alt="Preview">
                    <div class="avatar-inputs">
                        <input type="file" name="avatar" class="form-control" accept="image/*" onchange="previewAvatarFile(this)">
                        <input type="url" name="avatar_url" class="form-control" placeholder="Or Image URL" oninput="previewAvatarUrl(this.value)" value="<?= ($edit_data && strpos($edit_data[\'avatar\'], \'http\') === 0) ? htmlspecialchars($edit_data[\'avatar\'], ENT_QUOTES, \'UTF-8\') : \'\' ?>">
                    </div>
                </div>
            </div>

            <?php foreach ($SCHEMA_FIELDS as $field): ?>
            <div class="form-group">
                <label class="form-label"><?= htmlspecialchars($field, ENT_QUOTES, "UTF-8") ?></label>
                <input type="text" class="form-control" name="<?= htmlspecialchars($field, ENT_QUOTES, "UTF-8") ?>" value="<?= htmlspecialchars($edit_data[$field] ?? "", ENT_QUOTES, "UTF-8") ?>" required>
            </div>
            <?php endforeach; ?>

            <button type="submit" class="btn btn-primary"><?= $edit_data ? "Update Record" : "Save Record" ?></button>
            <?php if ($edit_data): ?>
                <a href="index.php" class="btn btn-secondary">Cancel</a>
            <?php endif; ?>
        </form>
    </div>

    <div class="card">
        <div class="card-header">
            <h2 class="card-title">All Records</h2>
            <a href="index.php?action=export" class="btn btn-excel">Export XLS</a>
        </div>
        
        <div class="table-container">
            <table>
                <thead>
                    <tr>
                        <th style="width: 50px;">Avatar</th>
                        <?php foreach ($SCHEMA_FIELDS as $field): ?>
                        <th><?= htmlspecialchars($field, ENT_QUOTES, "UTF-8") ?></th>
                        <?php endforeach; ?>
                        <th style="width: 110px;">Actions</th>
                    </tr>
                </thead>
                <tbody>
                    <?php if ($results): ?>
                        <?php while ($row = $results->fetchArray(SQLITE3_ASSOC)): ?>
                        <tr>
                            <td>
                                <?php if (!empty($row["avatar"])): ?>
                                    <img src="<?= htmlspecialchars($row["avatar"], ENT_QUOTES, "UTF-8") ?>" class="avatar-thumb">
                                <?php else: ?>
                                    <span style="color: var(--text-muted); font-size: 12px;">None</span>
                                <?php endif; ?>
                            </td>
                            <?php foreach ($SCHEMA_FIELDS as $field): ?>
                            <td><?= htmlspecialchars($row[$field] ?? "", ENT_QUOTES, "UTF-8") ?></td>
                            <?php endforeach; ?>
                            <td>
                                <div class="action-links">
                                    <a href="index.php?action=edit&id=<?= $row["id"] ?>" class="btn-action">Edit</a>
                                    <a href="index.php?action=delete&id=<?= $row["id"] ?>" class="btn-action btn-action-delete" onclick="return confirm(\'Delete this record?\')">Delete</a>
                                </div>
                            </td>
                        </tr>
                        <?php endwhile; ?>
                    <?php endif; ?>
                </tbody>
            </table>
        </div>
    </div>
</div>

<script>
function previewAvatarFile(input) {
    const preview = document.getElementById("avatar-preview");
    if (input.files && input.files[0]) {
        const reader = new FileReader();
        reader.onload = function(e) { preview.src = e.target.result; }
        reader.readAsDataURL(input.files[0]);
    }
}
function previewAvatarUrl(url) {
    const preview = document.getElementById("avatar-preview");
    if (url.trim() !== "") { preview.src = url; }
}
</script>

</body>
</html>';

    file_put_contents(__DIR__ . '/index.php', $index_code);

    header('Location: index.php');
    exit;
}
?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Dynamic System Builder</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        :root { --bg: #f8fafc; --surface: #ffffff; --border: #e2e8f0; --text-main: #0f172a; --text-muted: #64748b; --primary: #2563eb; --radius: 10px; --danger: #ef4444; }
        * { box-sizing: border-box; }
        body { font-family: 'Inter', -apple-system, sans-serif; background: var(--bg); color: var(--text-main); padding: 40px 20px; }
        .container { max-width: 720px; margin: 0 auto; background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius); padding: 32px; }
        h1 { font-size: 20px; font-weight: 700; margin: 0 0 24px 0; }
        .form-group { margin-bottom: 20px; }
        label { display: block; font-size: 13px; font-weight: 600; color: var(--text-muted); margin-bottom: 6px; }
        input[type="text"], input[type="file"], input[type="url"] { width: 100%; padding: 10px 12px; font-size: 14px; font-family: inherit; border: 1px solid var(--border); border-radius: 6px; background: #fff; }
        input:focus { outline: none; border-color: var(--primary); box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.1); }
        
        .field-input.has-error { border-color: var(--danger) !important; box-shadow: 0 0 0 3px rgba(239, 68, 68, 0.15) !important; }
        
        .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 24px; }
        .field-row { display: flex; flex-direction: column; gap: 4px; margin-bottom: 12px; }
        .field-row-inner { display: flex; gap: 8px; }
        .error-msg { color: var(--danger); font-size: 12px; font-weight: 500; display: none; }
        
        .btn { display: inline-flex; align-items: center; justify-content: center; padding: 10px 16px; font-size: 14px; font-weight: 500; border-radius: 6px; cursor: pointer; border: 1px solid transparent; }
        .btn-add { background: #f1f5f9; color: var(--text-main); border-color: var(--border); }
        .btn-remove { background: #fef2f2; color: #ef4444; border-color: #fca5a5; }
        .btn-submit { background: var(--primary); color: white; width: 100%; font-weight: 600; font-size: 15px; margin-top: 10px; }
        .btn-submit:disabled { background: #94a3b8; cursor: not-allowed; }
        .preview-img { max-height: 80px; margin-top: 10px; border-radius: 6px; display: none; }
    </style>
</head>
<body>
<div class="container">
    <h1>System Builder</h1>
    <form id="builder-form" method="POST" enctype="multipart/form-data">
        <input type="hidden" name="action_builder" value="1">
        
        <div class="form-group">
            <label>Project Name</label>
            <input type="text" name="system_name" placeholder="e.g. Employee Directory" required>
        </div>

        <div class="grid">
            <div>
                <label>System Logo (Optional)</label>
                <input type="file" name="system_logo" accept="image/*" onchange="previewImageFile(this)" style="margin-bottom: 8px;">
                <input type="url" name="system_logo_url" placeholder="Or Logo URL" oninput="previewImageUrl(this.value)">
                <img id="logo-preview" class="preview-img" alt="Logo Preview">
            </div>

            <div>
                <label>Database Fields</label>
                <div id="dynamic-fields">
                    <div class="field-row">
                        <div class="field-row-inner">
                            <input type="text" name="fields[]" class="field-input" placeholder="Field Name" oninput="validateFields()" required>
                            <button type="button" class="btn btn-remove" onclick="removeRow(this)">✕</button>
                        </div>
                        <span class="error-msg">This field name already exists!</span>
                    </div>
                </div>
                <button type="button" class="btn btn-add" onclick="addRow()">+ Add Field</button>
            </div>
        </div>

        <button type="submit" id="submit-btn" class="btn btn-submit">Build System</button>
    </form>
</div>

<script>
function previewImageFile(input) {
    const preview = document.getElementById('logo-preview');
    if (input.files && input.files[0]) {
        const reader = new FileReader();
        reader.onload = function(e) { preview.src = e.target.result; preview.style.display = 'block'; }
        reader.readAsDataURL(input.files[0]);
    }
}

function previewImageUrl(url) {
    const preview = document.getElementById('logo-preview');
    if (url.trim() !== '') { preview.src = url; preview.style.display = 'block'; }
}

function addRow() {
    const container = document.getElementById('dynamic-fields');
    const div = document.createElement('div');
    div.className = 'field-row';
    div.innerHTML = `
        <div class="field-row-inner">
            <input type="text" name="fields[]" class="field-input" placeholder="Field Name" oninput="validateFields()" required>
            <button type="button" class="btn btn-remove" onclick="removeRow(this)">✕</button>
        </div>
        <span class="error-msg">This field name already exists!</span>
    `;
    container.appendChild(div);
    validateFields();
}

function removeRow(btn) {
    const container = document.getElementById('dynamic-fields');
    if (container.querySelectorAll('.field-row').length > 1) { 
        btn.closest('.field-row').remove(); 
        validateFields();
    }
}

function validateFields() {
    const inputs = document.querySelectorAll('.field-input');
    const submitBtn = document.getElementById('submit-btn');
    const seen = new Map();
    let hasDuplicate = false;

    inputs.forEach(input => {
        const val = input.value.trim().toLowerCase();
        if (val !== '') {
            if (seen.has(val)) {
                seen.get(val).push(input);
            } else {
                seen.set(val, [input]);
            }
        }
    });

    inputs.forEach(input => {
        const row = input.closest('.field-row');
        const errorMsg = row.querySelector('.error-msg');
        const val = input.value.trim().toLowerCase();

        if (val !== '' && seen.get(val).length > 1) {
            input.classList.add('has-error');
            errorMsg.style.display = 'block';
            hasDuplicate = true;
        } else {
            input.classList.remove('has-error');
            errorMsg.style.display = 'none';
        }
    });

    submitBtn.disabled = hasDuplicate;
}
</script>
</body>
</html>