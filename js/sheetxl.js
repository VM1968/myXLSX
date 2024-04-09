let tMouse = {
    // isMouseDown
    // tMouse.target
    // tMouse.targetWidth
    // targetPosX
};
const eventNames = ["mousedown", "mouseup", "mousemove", "click"];
eventNames.forEach((e) => window.addEventListener(e, handle));

let active = null;

let range = null;
let rengebegin = null;
let fcell = null;
let ccell = null;
let xlrange = document.getElementById('xlrange');
let bcr1;
let bcr2;


function handle(e) {

    if (e.type === eventNames[0]) {
        //  console.log('mousedown');
        if (active) {
            active.classList.remove('active');
            let xlcorner = document.getElementById('xlcorner');
            xlcorner.style.left = '-2000px';
            xlcorner.style.top = '-2000px';
        }
        if (range) {
            range = null;
            xlrange.style.width = '0px';
            xlrange.style.height = '0px';
            xlrange.style.left = '-2000px';
            xlrange.style.top = '-2000px';
        }

        if (e.target.tagName == 'TD' && !e.target.classList.contains('row_selectall')) {
            rengebegin = e.target;
        }

        tMouse.isMouseDown = true;
        let element = e.target.parentElement;
        if (!element.dataset['td']) return false;
        let col = document.querySelector(`col[data-td='${element.dataset[`td`]}']`);
        let th = document.querySelector(`th[data-td='${element.dataset[`td`]}']`);

        tMouse.target = col;
        //console.log(tMouse);
        tMouse.targetWidth = col.clientWidth;
        tMouse.targetPosX = th.getBoundingClientRect().x;

        //console.log(tMouse);

    }
    if (e.type === eventNames[1]) {
        //  console.log('mouseup');
        tMouse = {};
        // console.log(range);
        if (range) {
            
            let bcr = xlrange.getBoundingClientRect();
            let allcell = document.querySelectorAll(`td`);
            let selected = [];
            let crow = [];
            let cY = bcr.y;
            allcell.forEach(cl => {
                bcr2 = cl.getBoundingClientRect();
                if (bcr2.x >= bcr.x && bcr2.x < bcr.x + bcr.width && bcr2.y >= bcr.y && bcr2.y < bcr.y + bcr.height) {
                    if (cY < bcr2.y) {
                        console.log(cY + ' ' + bcr2.y)
                        selected.push(crow);
                        crow = [];
                        cY = bcr2.y;
                    }
                    crow.push(cl.innerText);
                }

            })
            selected.push(crow);
            console.log(selected);

            MXLSX.SelectedRow=null;
        }
    };

    if (e.type === eventNames[2]) {
        // console.log('mousemove');
        
        if (tMouse.isMouseDown && (e.target.tagName == 'TD' && !e.target.classList.contains('row_selectall'))) {
            console.log(e.target);
            let ccell = e.target;
            if (!rengebegin) { rengebegin = ccell };
            if (ccell != fcell) {
                fcell = ccell;

                bcr1 = rengebegin.getBoundingClientRect();
                bcr2 = ccell.getBoundingClientRect();

                //console.log(bcr);
                let rtop = bcr1.y + window.pageYOffset;
                let rheight = bcr2.y + bcr2.height - bcr1.y;

                if (bcr2.y < rtop) {
                    rtop = bcr2.y + window.pageYOffset;
                    rheight = bcr1.y + bcr1.height - bcr2.y;
                }

                let rleft = bcr1.x + window.pageXOffset;
                let rwidth = bcr2.x + bcr2.width - bcr1.x;

                if (bcr2.x < rleft) {
                    rleft = bcr2.x + window.pageYOffset;
                    rwidth = bcr1.x + bcr1.width - bcr2.x;
                }

                xlrange.style.width = rwidth + 'px';
                xlrange.style.height = rheight + 'px';
                xlrange.style.left = rleft + 'px';
                xlrange.style.top = rtop + 'px';
            }
            range = true;
            MXLSX.SelectedRow=null;
        }

        if (tMouse.isMouseDown && e.target.id == 'xlrange') {
            if (e.clientX < bcr2.x || e.clientY < bcr2.y) {
                //console.log(e.clientX);
                xlrange.style.width = '0px';
                xlrange.style.height = '0px';
                xlrange.style.left = '-2000px';
                xlrange.style.top = '-2000px';
            }
        }


        if (!tMouse.target || !tMouse.isMouseDown) return false;
        let size = (e.clientX - tMouse.targetWidth) - tMouse.targetPosX;
        tMouse.target.width = tMouse.targetWidth + size + "px";


    }

    if (e.type === eventNames[3]) {
        //console.log('click');
        if (e.target.tagName == 'TD' && !e.target.classList.contains('row_selectall')) {

            let el = e.target;
            active = el;
            let bcr = active.getBoundingClientRect();
            //console.log(bcr);
            let ctop = bcr.y + bcr.height - 3 + window.pageYOffset;
            let cleft = bcr.x + bcr.width - 3 + window.pageXOffset;
            let xlcorner = document.getElementById('xlcorner');
            xlcorner.style.left = cleft + 'px';
            xlcorner.style.top = ctop + 'px';

            el.classList.add('active');
            MXLSX.SelectedRow=null;

        }

        if (e.target.tagName == 'TD' && e.target.classList.contains('row_selectall')) {
            //console.log('click row_selectall');
            let el = e.target;
            let cells = document.querySelectorAll(`td[data-tr=tr_${el.innerHTML}]`);
            let last = cells.length - 1;
            let bcr1 = cells[0].getBoundingClientRect();
            let bcr2 = cells[last].getBoundingClientRect();

            // console.log(bcr1);
            // console.log(bcr2);
            let rtop = bcr1.y + window.pageYOffset;
            let rleft = bcr1.x + window.pageXOffset;
            let rwidth = bcr2.x + bcr2.width - bcr1.x;//bcr1.width;
            let rheight = bcr1.height;
            let xlrange = document.getElementById('xlrange');
            xlrange.style.width = rwidth + 'px';
            xlrange.style.height = rheight + 'px';
            xlrange.style.left = rleft + 'px';
            xlrange.style.top = rtop + 'px';

            // console.log(xlrange);

            range = true;
            MXLSX.SelectedRow=el.innerHTML;
        }

        if (e.target.tagName == 'TH') {
            let el = e.target;
            let cells = document.querySelectorAll(`td[data-td='${el.dataset[`td`]}']`);
            //console.log(cells);
            let last = cells.length - 1;
            let bcr1 = cells[0].getBoundingClientRect();
            let bcr2 = cells[last].getBoundingClientRect();

            //console.log(bcr);
            let rtop = bcr1.y + window.pageYOffset;
            let rleft = bcr1.x + window.pageXOffset;
            let rwidth = bcr1.width;
            let rheight = bcr2.y + bcr2.height - bcr1.y;
            let xlrange = document.getElementById('xlrange');
            xlrange.style.width = rwidth + 'px';
            xlrange.style.height = rheight + 'px';
            xlrange.style.left = rleft + 'px';
            xlrange.style.top = rtop + 'px';

            // console.log(xlrange);

            range = true;
            MXLSX.SelectedRow=null;
        }

    }
}