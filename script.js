let table = document.querySelector(".sheet-body");
rows = document.querySelector(".rows");
columns = document.querySelector(".columns");
tableExists = false

const generateTable = () => {
    let rowsNumber = parseInt(rows.value), columnsNumber = parseInt(columns.value)
    if(isNaN(rowsNumber) || isNaN(columnsNumber)) {
        Swal.fire({
            icon: 'error',
            title: 'Oops...',
            text: 'You have to define Rows and Columns',
            confirmButtonText: 'Got It',
        })
    }
    table.innerHTML = ""
    for(let i=0; i<rowsNumber; i++){
        var tableRow = ""
        for(let j=0; j<columnsNumber; j++){
            tableRow += `<td contenteditable></td>`
        }
        table.innerHTML += tableRow
    }
    if(rowsNumber>0 && columnsNumber>0){
        tableExists = true
    }
}

const ExportToExcel = (type, fn, dl) => {
    let handler;
    if(!tableExists){
        clearTimeout(handler)
        Swal.fire({
            icon: 'error',
            title: 'Failed...',
            html: 'There is <strong style="color: tomato">NOTHING</strong> to Export '+'<br />'+'You Have To <strong style="color: tomato">Generate</strong> The Table First',
        })
    } else {
        var elt = table
        var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" })
        handler = setTimeout(() => {
            Swal.fire({
                position: 'top-end',
                icon: 'success',
                title: 'Your work has been saved',
                showConfirmButton: false,
                timer: 1500
            })
        }, 1500)
        return dl ? XLSX.write(wb, { bookType: type, bookSST: true, type: 'base64' })
            : XLSX.writeFile(wb, fn || ('MyNewSheet.' + (type || 'xlsx')))
    }
}