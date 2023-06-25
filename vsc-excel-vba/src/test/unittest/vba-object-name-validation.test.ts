import * as assert from "assert";
import { VBAObjectNameValidation } from "../../vba-object-name-validation";

suite("Test Suite VBA object name validation", () => {
    [
        ["test123", true], ["テスト123", true], 
        ["a".repeat(31), true], ["あ".repeat(15), true],
        ["a".repeat(15) + "あ".repeat(8), true],
        ["a".repeat(15) + "あ".repeat(8), true],
        ["a".repeat(32), false], ["あ".repeat(16), false],
        ["a".repeat(16) + "あ".repeat(8), false],
    ].forEach(param => {
        test(`validate len`, () => {
            const validate = new VBAObjectNameValidation();
            const text = param[0] as string;
            const result = param[1] as boolean;
            assert.equal(validate.len(text), result);
        });
    });

    [
        "test", "test1", "te1st", "test１", "te１st", "te_st", "te＿｡､ー｢｣・st"
    ].forEach(param => {
        test(`validate name ${param}`, () => {
            const validate = new VBAObjectNameValidation();
            assert.equal(validate.prefix(param), true);
            assert.equal(validate.symbol(param), true);
        });
    });

    [
        "_", "1", "1_", "_1", "123", "１"
    ].forEach(param => {
        test(`validate prefix ${param}`, () => {
            const input = `${param}test`;
            const validate = new VBAObjectNameValidation();
            assert.equal(validate.prefix(input), false);
        });
    });

    [
        " ", "　",
        "!",
        `"`,
        "#", "$", "%", "&", "'", "\"",
        "(", ")", "=", "-", "~", "^", "|", "\\", "`",
        "@", "{", "[", "+", ";", "*", ":", "}", "]", "<", ",", ">", ".", "?", "/",

        "！", "”", "＃", "＄", "％", "＆", "（", "）", 
        "＝", "～", "＾", "｜",   "＠", "｛"," ｝", "＋", "；", "＊", "；", 
        "：", "＊", "＜",  "＞", "？",

        "’", "￥","‘",
    ].forEach(param => {
        test(`validate synbol ${param}`, () => {
            const input = `te${param}st`;
            const validate = new VBAObjectNameValidation();
            assert.equal(validate.symbol(input), false);
        });
    });
});