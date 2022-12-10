<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">

    <title>E-Faktur generator</title>


    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0-beta.3/css/bootstrap.min.css" integrity="sha384-Zug+QiDoJOrZ5t4lssLdxGhVrurbmBWopoEl+M6BdEfwnCJZtKxi1KgxUyJq13dy" crossorigin="anonymous">

    <link rel="stylesheet" href="<?php echo base_url().'assets/css/jquery.dm-uploader.min.css'; ?>">
    <link rel="stylesheet" href="<?php echo base_url().'assets/css/main.css'; ?>">

  </head>

  <body>   

    <nav class="navbar navbar-expand-md navbar-dark bg-dark mb-5">
      <a class="navbar-brand" href="https://danielmg.org/">E-Faktur generator</a>
      <button class="navbar-toggler" type="button" data-toggle="collapse" data-target="#demosNavbarCollapse" aria-controls="demosNavbarCollapse" aria-expanded="false" aria-label="Toggle navigation">
        <span class="navbar-toggler-icon"></span>
      </button>
    </nav>

    <div class="container pt-2">
      <div class="row">
        <div class="col-md-6 col-sm-12">
        
          <div id="drag-and-drop-zone" class="dm-uploader p-5 text-center">
            <h3 class="mb-5 mt-5 text-muted">Drag &amp; drop Files Invoice Excel Hasil Export Kingdee disini</h3>

            <div class="btn btn-primary btn-block mb-5">
                <span>Pilih File</span>
                <input type="file" title="Click to add Files" multiple="">
            </div>
          </div><!-- /uploader -->

        </div>

        <div class="col-md-6 col-sm-12">
          <div class="card h-100">
            <div class="card-header">
              File List
            </div>

            <div class="card-content" style="min-height: 100px;">
              <ul class="list-unstyled p-2 d-flex flex-column col" id="files">
                <li class="text-muted text-center empty">No files uploaded.</li>
              </ul>
            </div>
          </div>
        </div>
      </div>

      <div class="row">
        <div class="col-12">
           <div class="card h-100">
            <div class="card-header">
              Process Messages
            </div>

            <ul class="list-group list-group-flush" id="debug">
            </ul>
          </div>
        </div>
      </div>

    </div>


    <script src="<?php echo base_url().'assets/js/jquery-3.6.0.min.js'; ?>"></script>
    <script src="<?php echo base_url().'assets/js/bootstrap.min.js'; ?>"></script>
    <script src="<?php echo base_url().'assets/js/jquery.dm-uploader.min.js'; ?>"></script>
    <script src="<?php echo base_url().'assets/js/upload-ui-log.js'; ?>"></script>

    <!-- File item template -->
    <script type="text/html" id="files-template">
      <li class="media">
        
        <div class="media-body mb-1">
          <p class="mb-2">
            <strong>%%filename%%</strong> <br/> Status: <span class="text-muted">Waiting</span>
          </p>
          <div class="progress mb-2">
            <div class="progress-bar progress-bar-striped progress-bar-animated bg-primary" 
              role="progressbar"
              style="width: 0%" 
              aria-valuenow="0" aria-valuemin="0" aria-valuemax="100">
            </div>
          </div>
          
          <hr class="mt-1 mb-1" />
        </div>
      </li>
    </script>

    <script type="text/html" id="debug-template">
      <li class="list-group-item text-%%color%%"><strong>%%date%%</strong>: %%message%%</li>
    </script>

    <script type="text/html" id="finish-template">
        <div class="col-sm-4">
          <figure class="figure">
            <img src="<?php echo base_url().'assets/images/csv.png'; ?>" class="figure-img img-fluid rounded" alt="result file xlsx" style="max-width: 50px;">
            <figcaption class="figure-caption">
              <strong>%%filename%%</strong><br/>
              <button type="button" class="btn btn-sm btn-primary" onclick="window.open('%%filepath%%');">Donwload</button>
            </figcaption>
          </figure>          
        </div>
      <hr class="mt-1 mb-1">
    </script>

    <script type="text/javascript">

      jQuery(document).ready(function() {
        let urlUpload = "<?php echo site_url('main/uploadFile')?>";

        let isProcessing = false;

        $("#drag-and-drop-zone").dmUploader({
          url: urlUpload,
          extFilter: ["xls", "xlsx"],
          
          onFileExtError : function(file) {
            alert('extensi file yang diperbolehkan hanya .xls dan .xlsx');
          },

          onDragEnter: function(){
            this.addClass('active');
          },
          onDragLeave: function(){
            this.removeClass('active');
          },
          onInit: function(){
            //ui_add_log('Penguin initialized :)', 'info');
          },
          onComplete: function(){
            // ui_add_log('All pending tranfers finished');
          },
          onNewFile: function(id, file){
            if (isProcessing) return false;
            else {
              isProcessing = true;
              ui_add_log('New file added #' + id);
              ui_multi_add_file(id, file);
            }
          },
          onBeforeUpload: function(id){
            ui_add_log('Starting the upload of #' + id);
            ui_multi_update_file_progress(id, 0, '', true);
            ui_multi_update_file_status(id, 'uploading', 'Uploading...');
          },
          onUploadProgress: function(id, percent){
            ui_multi_update_file_progress(id, percent);
          },
          onUploadSuccess: function(id, data){
            var jsonData = JSON.parse(data);
            ui_add_log('Server Response for file #' + id + ': ' + JSON.stringify(data));
            ui_add_log('Upload of file #' + id + ' COMPLETED', 'success');
            ui_multi_update_file_status(id, 'success', 'Upload Complete');
            // ui_multi_update_file_progress(id, 100, 'success', false);
            readDataFile(id, jsonData.file_name);
          },
          onUploadError: function(id, xhr, status, message){
            ui_multi_update_file_status(id, 'danger', message);
            ui_multi_update_file_progress(id, 0, 'danger', false);  
          },
          onFallbackMode: function(){
            ui_add_log('Plugin cant be used here, running Fallback callback', 'danger');
          },
          onFileSizeError: function(file){
            ui_add_log('File \'' + file.name + '\' cannot be added: size excess limit', 'danger');
          }
        });

        function readDataFile(id, fileName) {
          // console.log(fileName);

          ui_add_log('Processing data & generating e-faktur xlsx....');

          jQuery.ajax({
            url : "<?php echo site_url('main/readFile')?>",
            method: "POST",
            data: { file_name: fileName },
            dataType: "json",
            beforeSend : function(xhr) {
              ui_multi_update_file_status(id, 'success', 'Processing File');
            },
            success : function(res) {
              // console.log(res);
              if (res.success == 1) {
                ui_add_log('processing COMPLETED', 'success');
                ui_multi_update_file_status(id, 'success', 'Processing Finish');
                ui_multi_update_file_finish(id, res.filename, res.filepath);
                isProcessing = false;
              }
            },
            error : function() {
                ui_add_log('processing ERROR', 'danger');
            },
            complete : function() {
              isProcessing = false;
            },

          })
        }

      });




    </script>
  </body>
