import { dixonreed } from "../src/functions/functions";
import {test, expect} from "@jest/globals"

const data1 = [[0], [2], [3], [4]];
const data2 = [[0], [2], [3], [7]];
test("Dixon Reed Test", () => {
    let res = dixonreed(data1);
    expect(res[1][1]).toBeCloseTo(1/4);
    expect(res[2][1]).toBe(undefined);
    res = dixonreed(data2);
    expect(res[1][1]).toBeCloseTo(4/7);
    expect(res[2][1]).toBe(7);
});