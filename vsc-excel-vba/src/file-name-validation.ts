
export class FileNameValidation {
    prefix(objectName: string): boolean {
        const name = this.replaceZenkakuSymbol(objectName);
        const reg = /^[\d_０-９]/;
        return !reg.test(name);
    }
    
    symbol(objectName: string): boolean {
        const name = this.replaceZenkakuSymbol(objectName);
        const reg = /[\s　!"#$%&'\(\)=\-~^|\\`@{\[\+;\*:}\]<,>.?\/”’￥‘]/g;
        return !reg.test(name);
    }
    
    len(objectName: string): boolean {
        let len = 0;
        for (const char of [...objectName]) {
            if (char.match(/[ -~]/) !== null) {
                len += 1;
            } else {
                len += 2;
            }
        }
        return len < 32;
    }

    private replaceZenkakuSymbol(objectName: string): string {
        return objectName.replace(/[！＃＄％＆（）＝ー～＾｜＠｛｝＋；＊：＜、＞。？・]/g, x => {
            return String.fromCharCode(x.charCodeAt(0)-0xFEE0);
        });
    } 
}
