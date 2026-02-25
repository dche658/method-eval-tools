/* Range input component.
*
* If numbers are allowed then accept numbers or an Excel range specification.
* If numbers are not allowed, then only accept an Excel range specification.
* 
* A range is specified in the form A3:B23 or as an absolute reference using
* the dollar symbol, e.g. $A$3:$B:$23. Various combinations may be used.
* 
* The sheet name may also be specified. e.g. Sheet1!A3:B23. This is the form
* that is used if the range is read from the selected cells.
* 
* Author: Douglas Chesher
*
* Created: October 2025.
*/
import * as React from "react";
import { useState } from "react";
import { Button, Field, Input, Tooltip, makeStyles } from "@fluentui/react-components";
import { AppsRegular } from "@fluentui/react-icons";


// ^(?:([a-zA-Z\\d\\x5F\\x2D]*\\x21)? Begin optionally with number, letter, hypen, or underscore followed by exclamation mark.
// (\\$?[A-Z]{1,3}\\$?[1-9]{1}\\d{0,6}) Optionally match dollar sign followed by 1-3 letters followed optionally by dollar sign followed by one or more digits greater than 1
// (:\\$?[A-Z]{1,3}\\$?[1-9]{1}\\d{0,6})?)$ Optionally match group beginning with colon followed optionally by dollar sign followed by 1-3 letters followed optionally by dollar sign followed by one or more digits greater than 1.
const REGEX_RANGE_ONLY = "^(?:([a-zA-Z\\d\\x5F\\x2D]*\\x21)?(\\$?[A-Z]{1,3}\\$?[1-9]{1}\\d{0,6})(:\\$?[A-Z]{1,3}\\$?[1-9]{1}\\d{0,6})?)$"

// Optionally match integer or decimal number or as above.
const REGEX_RANGE_PLUS_NUMBERS = "^(?:(\\d+\\x2E?\\d*)?(([a-zA-Z\\d\\x5F\\x2D]*\\x21)?(\\$?[A-Z]{1,3}\\$?[1-9]{1}\\d{0,6})(:\\$?[A-Z]{1,3}\\$?[1-9]{1}\\d{0,6})?)?)$"

interface RangeInputProps {
    rangeValue: string;
    setRangeValue?: (value: string) => void;
    label: string;
    validationMessage?: string;
    tooltipContent?: string;
    allowNumbers?: boolean;
    uitext: {[key:string]: string};
}

const useStyles = makeStyles({
    root: {
        display: "flex",
        alignItems: "center",
        columnGap: "4px",
    },
    field: {
        marginLeft: "4px",
    },
});

const RangeInput: React.FC<RangeInputProps> = (props: RangeInputProps) => {
    const [validationMessage, setValidationMessage] = useState<string>("");

    const styles = useStyles();

    const selectRange = async () => {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load("address");
            await context.sync();
            //console.log(`Selected range address: ${range.address}`);
            props.setRangeValue(range.address);
        });
    }

    const handleTextChange = async (event: React.ChangeEvent<HTMLInputElement>) => {
        props.setRangeValue(event.target.value);
        if (event.target.validity.patternMismatch) {
            setValidationMessage(props.validationMessage || "Invalid format.");
        } else {
            setValidationMessage("");
        }
    };

    function hasTooltip() {
        const PATTERN = props.allowNumbers ? REGEX_RANGE_PLUS_NUMBERS : REGEX_RANGE_ONLY;
        if (props.tooltipContent) {
            return (
                <Tooltip content={props.tooltipContent} relationship="label">
                    <Input
                        value={props.rangeValue}
                        className="mb-3"
                        pattern={PATTERN}
                        onChange={handleTextChange}
                    />
                </Tooltip>
            );
        } else {
            return (
                <Input
                    value={props.rangeValue}
                    className="mb-3"
                    pattern={PATTERN}
                    onChange={handleTextChange}
                />
            );
        }
    }

    return (
        <div className={styles.field}>
            <Field label={props.label}
                validationMessage={validationMessage}>
                <div className={styles.root} >
                    {hasTooltip()}
                    <Tooltip content={props.uitext["tip_range_copy"]} relationship="label">
                        <Button icon={<AppsRegular />} onClick={selectRange} />
                    </Tooltip>
                </div>
            </Field>
        </div>
    );
};

export default RangeInput;