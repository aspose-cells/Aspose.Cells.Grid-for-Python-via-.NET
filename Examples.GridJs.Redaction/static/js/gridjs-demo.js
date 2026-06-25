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

    const zorder = 1000;               // Canvas Zorder
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
                // UI language (change as needed)
                local: 'en',
                mode: 'read',
                enableRedactionShape: true,
                redactionDefaultColor: 'green',
                redactionReasons: [
            	'Personal Information',
            	'Confidential',
            	'Secret',
                ],
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
        const activeSheetName = jsondata.actname;
        // Initialise GridJs (x_spreadsheet) and bind to the container,load data and set active sheet
        xs = x_spreadsheet('#gridjs-demo-uid', option)
            .loadData(sheets, activeSheetName)
            .updateCellError(msg => console.error(msg));

        // Hide the bottom sheet bar when tabs are disabled
        if (!jsondata.showtabs) {
            xs.bottombar.hide();
        }

        xs.setUniqueId(jsondata.uniqueid);
        xs.setFileName(filename);
    // ----------------------------------------------------------------
    //  4. Set the active cell (fallback to A1 if needed)
    // ----------------------------------------------------------------

    if (xs.getActiveSheetName() == activeSheetName)  {
        xs.setActiveCell(jsondata.actrow, jsondata.actcol);
        } else {
        // Fallback – set A1 as active cell
        xs.setActiveCell(0, 0);
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
        xs.on('redaction-inserted', (s, r) =>
    	{
        	console.log('redaction-inserted:', s, " item:", r);

    	}).on('redaction-updated', (s, r) =>
    	{
        	console.log('redaction-updated:', s, " item:", r);

    	}).on('redaction-deleted', (s, r) =>
    	{
        	console.log('redaction-deleted:', s, " item:", r);

    	});
    }

})(jQuery);
