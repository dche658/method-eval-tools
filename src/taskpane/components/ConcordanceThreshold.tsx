import * as React from "react";

import { Input, makeStyles } from "@fluentui/react-components";
import { tokens } from "@fluentui/react-components"

interface ConcordanceThresholdProps {
    xThreshold0: string;
    setXThreshold0: React.Dispatch<React.SetStateAction<string>>;
    xThreshold1: string;
    setXThreshold1: React.Dispatch<React.SetStateAction<string>>;
    xThreshold2: string;
    setXThreshold2: React.Dispatch<React.SetStateAction<string>>;
    xThreshold3: string;
    setXThreshold3: React.Dispatch<React.SetStateAction<string>>;
    yThreshold0: string;
    setYThreshold0: React.Dispatch<React.SetStateAction<string>>;
    yThreshold1: string;
    setYThreshold1: React.Dispatch<React.SetStateAction<string>>;
    yThreshold2: string;
    setYThreshold2: React.Dispatch<React.SetStateAction<string>>;
    yThreshold3: string;
    setYThreshold3: React.Dispatch<React.SetStateAction<string>>;
}

const useStyles = makeStyles({
    container: {
        display: "flex",
    },
    gridcontainer: {
        display: "grid",
        gridTemplateColumns: "120px 120px",
        backgroundColor: tokens.colorNeutralCardBackground,
        borderRight: "1px solid "+tokens.colorPaletteAnchorBorderActive,
        borderBottom: "1px solid "+tokens.colorPaletteAnchorBorderActive,
    },
    griddiv: {
        borderTop: "1px solid "+tokens.colorPaletteAnchorBorderActive,
        borderLeft: "1px solid "+tokens.colorPaletteAnchorBorderActive,
        padding: "3px",
        justifyContent: "center",
        alignItems: "center",
        display: "flex",
        textAlign: "center",
    },
    inputfield: {
        width: "100%",
    }
});

export default function ConcordanceThreshold(props: ConcordanceThresholdProps) {
    const styles = useStyles();

    const handleXThreshold0Change = (event: React.ChangeEvent<HTMLInputElement>) => {
        props.setXThreshold0(event.target.value);
    };
    const handleXThreshold1Change = (event: React.ChangeEvent<HTMLInputElement>) => {
        props.setXThreshold1(event.target.value);
    };
    const handleXThreshold2Change = (event: React.ChangeEvent<HTMLInputElement>) => {
        props.setXThreshold2(event.target.value);
    };
    const handleXThreshold3Change = (event: React.ChangeEvent<HTMLInputElement>) => {
        props.setXThreshold3(event.target.value);
    };
    const handleYThreshold0Change = (event: React.ChangeEvent<HTMLInputElement>) => {
        props.setYThreshold0(event.target.value);
    };
    const handleYThreshold1Change = (event: React.ChangeEvent<HTMLInputElement>) => {
        props.setYThreshold1(event.target.value);
    };
    const handleYThreshold2Change = (event: React.ChangeEvent<HTMLInputElement>) => {
        props.setYThreshold2(event.target.value);
    };
    const handleYThreshold3Change = (event: React.ChangeEvent<HTMLInputElement>) => {
        props.setYThreshold3(event.target.value);
    };

    return (
        <div className={styles.container}>
            <div className={styles.gridcontainer}>
                <div className={styles.griddiv}>
                    <label>X CutOff</label>
                </div>
                <div className={styles.griddiv}>
                    <label>Y CutOff</label>
                </div>
                <div className={styles.griddiv}>
                    <Input value={props.xThreshold0} onChange={handleXThreshold0Change}
                        className={styles.inputfield}
                        type="number" />
                </div>
                <div className={styles.griddiv}>
                    <Input value={props.yThreshold0} onChange={handleYThreshold0Change}
                        className={styles.inputfield}
                        type="number" />
                </div>
                <div className={styles.griddiv}>
                    <Input value={props.xThreshold1} onChange={handleXThreshold1Change}
                        className={styles.inputfield}
                        type="number" />
                </div>
                <div className={styles.griddiv}>
                    <Input value={props.yThreshold1} onChange={handleYThreshold1Change}
                        className={styles.inputfield}
                        type="number" />
                </div>
                <div className={styles.griddiv}>
                    <Input value={props.xThreshold2} onChange={handleXThreshold2Change}
                        className={styles.inputfield}
                        type="number" />
                </div>
                <div className={styles.griddiv}>
                    <Input value={props.yThreshold2} onChange={handleYThreshold2Change}
                        className={styles.inputfield}
                        type="number" />
                </div>
                <div className={styles.griddiv}>
                    <Input value={props.xThreshold3} onChange={handleXThreshold3Change}
                        className={styles.inputfield}
                        type="number" />
                </div>
                <div className={styles.griddiv}>
                    <Input value={props.yThreshold3} onChange={handleYThreshold3Change}
                        className={styles.inputfield}
                        type="number" />
                </div>
            </div>
        </div>

    );
}