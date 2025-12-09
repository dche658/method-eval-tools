import * as React from "react";
import { useState } from "react";
import { Button, Field, Input, Tooltip, tokens, makeStyles } from "@fluentui/react-components";
import { AppsRegular } from "@fluentui/react-icons";

interface RangeInputProps {
    rangeValue: string;
    setRangeValue?: (value: string) => void;
    label: string;
    validationMessage?: string;
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
            console.log(`Selected range address: ${range.address}`);
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
    return (
        <div className={styles.field}>
        <Field label={props.label}
        validationMessage={validationMessage}>
            <div className={styles.root} >
            <Input
                value={props.rangeValue}
                className="mb-3"
                pattern="^(?:([a-zA-Z\d\x5F\x2D]*\x21)?(\$?[A-Z]{1,3}\$?[1-9]{1}\d{0,6})(:\$?[A-Z]{1,3}\$?[1-9]{1}\d{0,6})?)$"
                onChange={handleTextChange}
            />
            <Tooltip content="Copy selected range" relationship="label">
                <Button icon={<AppsRegular />} onClick={selectRange} />
            </Tooltip>
            </div>
        </Field>
        </div>
    );
};

export default RangeInput;