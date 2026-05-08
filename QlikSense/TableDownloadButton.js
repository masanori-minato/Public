/**
 * TableDownloadButton.js
 * ストレートテーブルのデータをCSVダウンロードするボタン拡張
 *
 * 使い方:
 *   1. Qlik Sense のシートに本エクステンションを配置する
 *   2. プロパティパネルで「テーブルオブジェクトID」にダウンロードしたい
 *      ストレートテーブルのオブジェクトID（例: ABCdEf）を入力する
 *   3. ボタンをクリックするとCSVがダウンロードされる
 */
define(["qlik", "jquery"], function (qlik, $) {
  "use strict";

  /* ------------------------------------------------------------------ */
  /* ユーティリティ                                                        */
  /* ------------------------------------------------------------------ */

  /**
   * 文字列をCSVセル用にエスケープする（ダブルクォートで囲み " を "" に置換）
   */
  function escapeCell(value) {
    var str = (value == null) ? "" : String(value);
    return '"' + str.replace(/"/g, '""') + '"';
  }

  /**
   * 行データ配列からCSV文字列を生成する
   * @param {string[]} headers  - ヘッダー文字列配列
   * @param {Array[]}  rows     - セルオブジェクト配列の配列 (qMatrix 形式)
   * @param {string}   sep      - 区切り文字
   * @param {boolean}  withHeader
   * @returns {string}
   */
  function buildCsv(headers, rows, sep, withHeader) {
    var lines = [];

    if (withHeader) {
      lines.push(headers.map(escapeCell).join(sep));
    }

    rows.forEach(function (row) {
      var cells = row.map(function (cell) {
        // qIsNull: NULL値は空文字
        return escapeCell(cell.qIsNull ? "" : cell.qText);
      });
      lines.push(cells.join(sep));
    });

    // UTF-8 BOM を先頭に付けて Excel でも文字化けしないようにする
    return "\uFEFF" + lines.join("\n");
  }

  /**
   * ブラウザのダウンロードダイアログを起動する
   */
  function triggerDownload(csvString, fileName) {
    var blob = new Blob([csvString], { type: "text/csv;charset=utf-8;" });
    var url  = URL.createObjectURL(blob);
    var a    = document.createElement("a");
    a.href     = url;
    a.download = fileName.replace(/\.csv$/i, "") + ".csv";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 500);
  }

  /**
   * ハイパーキューブの全ページを再帰的に取得する
   * @param {object}   model     - Qlik オブジェクトモデル
   * @param {number}   totalRows
   * @param {number}   totalCols
   * @param {number}   pageSize  - 1回に取得する行数（最大 10000）
   * @returns {Promise<Array[]>} - 全行の配列
   */
  function fetchAllRows(model, totalRows, totalCols, pageSize) {
    var pages = [];
    for (var offset = 0; offset < totalRows; offset += pageSize) {
      pages.push({
        qLeft  : 0,
        qTop   : offset,
        qWidth : totalCols,
        qHeight: Math.min(pageSize, totalRows - offset)
      });
    }

    var fetches = pages.map(function (page) {
      return model.getHyperCubeData("/qHyperCubeDef", [page]);
    });

    return qlik.Promise.all(fetches).then(function (results) {
      var allRows = [];
      results.forEach(function (pageResult) {
        pageResult[0].qMatrix.forEach(function (row) {
          allRows.push(row);
        });
      });
      return allRows;
    });
  }

  /* ------------------------------------------------------------------ */
  /* メイン処理：クリック時にCSVを生成してダウンロード                      */
  /* ------------------------------------------------------------------ */

  function runDownload(app, objectId, fileName, separator, includeHeader, $btn, originalLabel, originalColor) {
    $btn.text("読み込み中…").prop("disabled", true).css("cursor", "wait");

    app.getObject(objectId)
      .then(function (model) {
        return model.getLayout().then(function (layout) {
          var hc = layout.qHyperCube;

          if (!hc) {
            throw new Error("指定されたオブジェクトにハイパーキューブが見つかりません。\nストレートテーブルのオブジェクトIDを指定してください。");
          }

          // ヘッダー構築（ディメンション → メジャーの順）
          var dims    = hc.qDimensionInfo || [];
          var meas    = hc.qMeasureInfo   || [];
          var headers = []
            .concat(dims.map(function (d) { return d.qFallbackTitle; }))
            .concat(meas.map(function (m) { return m.qFallbackTitle; }));

          var totalRows = hc.qSize.qcy;
          var totalCols = hc.qSize.qcx;

          if (totalRows === 0) {
            throw new Error("テーブルにデータがありません。");
          }

          $btn.text("取得中… (0 / " + totalRows + " 行)");

          return fetchAllRows(model, totalRows, totalCols, 2000)
            .then(function (allRows) {
              return { headers: headers, rows: allRows };
            });
        });
      })
      .then(function (result) {
        var csv = buildCsv(result.headers, result.rows, separator, includeHeader);
        triggerDownload(csv, fileName);
        $btn.text("✓ ダウンロード完了").css({ background: "#4caf50", cursor: "pointer" });
        setTimeout(function () {
          $btn.text(originalLabel).css({ background: originalColor, cursor: "pointer" }).prop("disabled", false);
        }, 2000);
      })
      .catch(function (err) {
        console.error("[TableDownloadButton] Error:", err);
        alert("エラー: " + (err.message || err));
        $btn.text(originalLabel).css({ background: originalColor, cursor: "pointer" }).prop("disabled", false);
      });
  }

  /* ------------------------------------------------------------------ */
  /* エクステンション定義                                                  */
  /* ------------------------------------------------------------------ */

  return {
    /* ---------- プロパティパネル ---------- */
    definition: {
      type     : "items",
      component: "accordion",
      items    : {
        downloadSettings: {
          type : "items",
          label: "ダウンロード設定",
          items: {
            tableObjectId: {
              ref         : "props.tableObjectId",
              label       : "テーブルオブジェクトID ※必須",
              type        : "string",
              defaultValue: "",
              expression  : "none"
            },
            fileName: {
              ref         : "props.fileName",
              label       : "ダウンロードファイル名（拡張子なし）",
              type        : "string",
              defaultValue: "export",
              expression  : "none"
            },
            separator: {
              ref      : "props.separator",
              label    : "区切り文字",
              type     : "string",
              component: "dropdown",
              options  : [
                { value: ",",  label: "カンマ ( , )" },
                { value: "\t", label: "タブ"         },
                { value: ";",  label: "セミコロン ( ; )" }
              ],
              defaultValue: ","
            },
            includeHeader: {
              ref         : "props.includeHeader",
              label       : "ヘッダー行を含める",
              type        : "boolean",
              defaultValue: true
            }
          }
        },
        buttonSettings: {
          type : "items",
          label: "ボタン設定",
          items: {
            buttonLabel: {
              ref         : "props.buttonLabel",
              label       : "ボタンのラベル",
              type        : "string",
              defaultValue: "📥 CSVダウンロード",
              expression  : "optional"
            },
            buttonColor: {
              ref         : "props.buttonColor",
              label       : "ボタンの背景色",
              type        : "string",
              defaultValue: "#3F8AB3",
              expression  : "none"
            },
            buttonTextColor: {
              ref         : "props.buttonTextColor",
              label       : "ボタンの文字色",
              type        : "string",
              defaultValue: "#ffffff",
              expression  : "none"
            },
            fontSize: {
              ref         : "props.fontSize",
              label       : "フォントサイズ (px)",
              type        : "number",
              defaultValue: 14
            }
          }
        }
      }
    },

    initialProperties: {
      version: 1
    },

    /* ---------- 描画 ---------- */
    paint: function ($element, layout) {
      var me    = this;
      var props = layout.props || {};

      // プロパティ取得
      var tableObjectId  = (props.tableObjectId  || "").trim();
      var fileName       = (props.fileName       || "export").trim() || "export";
      var separator      = props.separator       !== undefined ? props.separator : ",";
      var includeHeader  = props.includeHeader   !== false;
      var buttonLabel    = (props.buttonLabel    || "📥 CSVダウンロード");
      var buttonColor    = (props.buttonColor    || "#3F8AB3");
      var buttonTextColor= (props.buttonTextColor|| "#ffffff");
      var fontSize       = (props.fontSize       || 14);

      // コンテナをリセット
      $element.empty().css({
        display       : "flex",
        alignItems    : "center",
        justifyContent: "center",
        width         : "100%",
        height        : "100%"
      });

      var isConfigured = tableObjectId.length > 0;

      var $btn = $("<button>")
        .text(isConfigured ? buttonLabel : "⚠ テーブルIDを設定してください")
        .css({
          padding      : "10px 24px",
          background   : isConfigured ? buttonColor : "#999999",
          color        : buttonTextColor,
          border       : "none",
          borderRadius : "4px",
          cursor       : isConfigured ? "pointer" : "not-allowed",
          fontSize     : fontSize + "px",
          fontFamily   : "inherit",
          fontWeight   : "500",
          lineHeight   : "1.4",
          transition   : "opacity 0.15s ease",
          maxWidth     : "100%",
          wordBreak    : "break-word"
        })
        .on("mouseenter", function () {
          if (isConfigured) { $(this).css("opacity", "0.85"); }
        })
        .on("mouseleave", function () {
          $(this).css("opacity", "1"); }
        );

      if (isConfigured) {
        $btn.on("click", function () {
          var app = qlik.currApp(me);
          runDownload(app, tableObjectId, fileName, separator, includeHeader, $btn, buttonLabel, buttonColor);
        });
      }

      $element.append($btn);

      return qlik.Promise.resolve();
    }
  };
});
