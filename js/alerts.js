(function($) {
  showSwal = function(type) {
    'use strict';
    if (type === 'custom-html') {
      swal({
        content: {
          element: "input",
          attributes: {
            placeholder: "CIAOOO",
            type: "password",
            class: 'form-control'
          },
        },
        buttons: {
          cancel: {
            text: "Cancel",
            value: null,
            visible: true,
            className: "btn btn-danger",
            closeModal: true,
          },
          confirm: {
            text: "OK",
            value: true,
            visible: true,
            className: "btn btn-primary",
            closeModal: true
          }
        }
      })
    } 
  }

})(jQuery);