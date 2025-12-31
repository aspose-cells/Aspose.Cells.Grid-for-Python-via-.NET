/* static/js/gridjs-demo.js */
/* -------------------------------------------------------------
   Client‑side script that initializes GridJs, loads the workbook,
   and wires up all server‑side URLs.
   ------------------------------------------------------------- */

(function ($) {
    // ---------- 1. Helper: generate a UUID ----------
    function generateUUID() {
        // Simple RFC4122 version 4 UUID
        return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
            var r = Math.random() * 16 | 0,
                v = c === 'x' ? r : (r & 0x3 | 0x8);
            return v.toString(16);
        });
    }

    // ---------- 2. Define server URLs ----------
    const queryJsonUrl      = "/GridJs/LoadSpreadsheet";
    const updateUrl         = "/GridJs/UpdateCell";
    const fileDownloadUrl   = "/GridJs/Download";
    const oleDownloadUrl    = "/GridJs/Ole";
    const imageurl          = "/GridJs/ImageUrl";
    const imageuploadurl1   = "/GridJs/AddImage";
    const imageuploadurl2   = "/GridJs/AddImageByURL";
    const imagecopyurl      = "/GridJs/CopyImage";

    const zorder = 1000;               // Canvas Z‑order
    let xs;                            // GridJs instance
    const uid = generateUUID();        // Unique session id

    // ---------- 3. Load workbook JSON ----------
    $.ajax({
        url: queryJsonUrl,
        method: "GET",
        data: { uid: uid },
        dataType: "json",
        // The server returns gzip‑compressed JSON; browsers decompress automatically.
        success: function (responseData) {
            const jsondata = typeof responseData === "string"
                ? JSON.parse(responseData)
                : responseData;

            const option = {
                updateMode: "server",
                updateUrl: updateUrl,
                local: "en"
            };

            loadWithOption(jsondata, option);
        },
        error: function (xhr) {
            console.error("Failed to load spreadsheet:", xhr);
        }
    });

    // ---------- 4. Render the UI ----------
    function loadWithOption(jsondata, option) {
        $('#gridjs-demo-uid').empty();

        const sheets = jsondata.data;
        const filename = jsondata.filename;

        // Initialise GridJs (x_spreadsheet) and bind to the container
        xs = x_spreadsheet('#gridjs-demo-uid', option)
            .loadData(sheets)
            .updateCellError(msg => console.error(msg));

        // Hide the bottom sheet bar when tabs are disabled
        if (!jsondata.showtabs) {
            xs.bottombar.hide();
        }

        xs.setUniqueId(jsondata.uniqueid);
        xs.setFileName(filename);

        // Activate the appropriate sheet & cell
        let activeSheetName = jsondata.actname;
        if (xs.bottombar.dataNames.includes(activeSheetName)) {
            xs.setActiveSheetByName(activeSheetName)
              .setActiveCell(jsondata.actrow, jsondata.actcol);
        } else {
            // Fallback to the first visible sheet
            activeSheetName = xs.bottombar.dataNames[0];
            xs.setActiveSheetByName(activeSheetName).setActiveCell(0, 0);
        }

        // ---------- 5. Register auxiliary URLs ----------
        xs.setImageInfo(
            imageurl,
            imageuploadurl1,
            imageuploadurl2,
            imagecopyurl,
            zorder,
            "/image/loading.gif"
        );
        xs.setFileDownloadInfo(fileDownloadUrl);
        xs.setOleDownloadInfo(oleDownloadUrl);
        xs.setOpenFileUrl("/GridJs/Index");
    }

})(jQuery);
