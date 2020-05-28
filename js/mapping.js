var docLocation = document.getElementById('doc_location');

docLocation.onchange = function(){
    document.getElementById('doc_subj_of_juri').value = docLocation.value;
}

var docDay = document.getElementById('doc_day');

docDay.onchange = function() {
    document.getElementById('doc_executed_day').value = docDay.value;
}