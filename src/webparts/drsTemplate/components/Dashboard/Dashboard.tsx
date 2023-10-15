
import * as React from 'react';

import { Link } from "react-router-dom";
import { IDashboardProps } from "./IDashboardProps";
import MUIDataTable from "mui-datatables";
import styles from './Dashboard.module.scss';
import * as moment from 'moment';
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import CircularProgress from '@material-ui/core/CircularProgress';
import { Folder, FileCopy, Share, AttachFile, CheckCircleOutline, Info, Delete, DoneAllOutlined, CloseOutlined, HourglassEmptyOutlined } from '@material-ui/icons';
import { TableContainer, Table, TableBody, TableCell, TableHead, TableRow, Paper, Tooltip, Backdrop, makeStyles, createStyles, Theme } from '@material-ui/core';
import { CopyToClipboard } from 'react-copy-to-clipboard';
import Dropzone from 'react-dropzone';
import { Dialog, DialogFooter, DefaultButton, PrimaryButton, DialogType, Label, Panel, PanelType, TextField } from 'office-ui-fabric-react';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { DisplayMode } from '@microsoft/sp-core-library';

// import { sp } from '@pnp/sp';
// import '@pnp/sp/webs';
// import '@pnp/sp/lists';
// import '@pnp/sp/items';

//let _queryString: any = require('query-string');

const today: Date = new Date(Date.now());


const controlClass = mergeStyleSets({
    control: {
        margin: '0 0 15px 0',
        maxWidth: '300px',
    },
});

let _filter = "";

const onFormatDate = (date: Date): string => {
    return (date.getMonth() + 1) + '/' + date.getDate() + '/' + (date.getFullYear() % 100);
};

const modelProps = {
    isBlocking: true,
    topOffsetFixed: true,
    styles: { main: { minWidth: 750 } },
};

const dialogContentProps = {
    type: DialogType.largeHeader,
    title: 'ADDITIONAL ATTACHMENTS (IF NEEDED)',
    subText: '',
};

const useStyles = makeStyles((theme: Theme) =>
    createStyles({
        root: {
            "& .MuiTextField-root": {
                margin: theme.spacing(1),
                width: 200
            },
            "& .MuiFormControl-root": {
                width: 100
            },

            button: {
                marginTop: theme.spacing(1),
                marginRight: theme.spacing(1)
            },
            actionsContainer: {
                marginBottom: theme.spacing(2)
            },
            resetContainer: {
                padding: theme.spacing(3)
            }

        }
    })
);

const backdropStyles = makeStyles((theme: Theme) =>
    createStyles({
        backdrop: {
            zIndex: theme.zIndex.drawer + 1,
            color: '#fff',
        },
    }),
);

export class Dashboard extends React.Component<IDashboardProps, any> {
    private today = new Date(Date.now());
    private gridColumns = [];
    private isFiterRequire = false;

    public constructor(props: IDashboardProps) {
        super(props);

        this.state = {
            items: [],
            AllDocuments: [],
            Showfilter: false,
            loading: true,
            currentPage: 1,
            searchStartDate: "",
            searchEndDate: "",
            searchRequestitems: [],
            hideDateRangeModal: true,
            dateErrorMsg: "",
            hideDialog: true, Approvers: {}, acceptedFiles: [], fileMetadata: [], selectedDRS: "", selectedFileID: "",
            hideApprovalPane: true, saveApproval: false,
            RecommendApprovers: null, AgreeApprovers: [], PerformApprovers: null, InputApprovers: [], DecisionApprovers: null
        };
        // _filter = this.props.filter.toString();
        //let queryParm = _queryString.parse(window.location.hash.substring(2));

        this.gridColumns = [
            {
                name: "Id",
                label: "Req #",
                download: false,
                options: {
                    download: false,
                    filter: false,
                    display: false,
                    viewColumns: false,
                    sort: true
                }
            },
            {
                name: "fileURL",
                label: "fileURL #",
                download: false,
                options: {
                    download: false,
                    filter: false,
                    display: false,
                    viewColumns: false,
                    sort: false
                }
            },
            {
                name: "DRSName",
                label: "DRSName",
                download: false,
                options: {
                    filter: true,

                    sort: true,
                    customBodyRender: (value, t, f, v) => {
                        let navLink = t.rowData[1];
                        return (
                            <a href={navLink}>{value}</a>


                        );
                    }
                }
            },
            {
                name: "DRSNumber",
                label: "DRSNumber",
                download: true,
                options: {
                    filter: false,
                    sort: true,
                }
            }, {
                name: "DRSStatus",
                label: "DRS Status",
                options: {
                    filter: true,
                    //filterList: queryParm["Status"] == undefined ? null : [queryParm["Status"]], ParentStatus
                    sort: true,
                }
            },
            {
                name: "ProjectName",
                label: "Project Name",
                download: false,
                options: {
                    filter: true,
                    //filterList: queryParm["Status"] == undefined ? null : [queryParm["Status"]], ParentStatus
                    sort: true,
                }
            },

            {
                name: "Project_Number",
                label: "Project Number",
                download: false,
                options: {
                    filter: true,
                    //filterList: queryParm["Status"] == undefined ? null : [queryParm["Status"]],TotalRefund
                    sort: true,
                }
            },
            {
                name: "TargetDate",
                label: "Date Created",
                download: false,
                options: {
                    filter: true,
                    sort: true,
                    customBodyRender: (value, t, f, v) => {
                        let navLink = this.props.folderPath;
                        return (
                            moment(value).format("DD/MM/YYYY")

                        );
                    }
                }
            },
            {
                name: "",
                label: "CopyLink",
                download: false,
                options: {
                    filter: false,
                    sort: false,
                    customBodyRender: (value, t, f, v) => {
                        let navLink = this.props.folderPath;
                        return (
                            <CopyToClipboard onCopy={() => {
                                alert("Link Copied Successfully");
                            }} text={"https://beachenergy.sharepoint.com" + encodeURI(t.rowData[1])}>
                                <FileCopy style={{ color: "#64c0eb" }}></FileCopy>
                            </CopyToClipboard>

                        );
                    }
                }
            },
            // {
            //     name: "",
            //     label: "",
            //     download: false,
            //     options: {
            //         filter: false,
            //         sort: false,
            //         customBodyRender: (value, t, f, v) => {
            //             let navLink = this.props.folderPath;
            //             return (
            //                 <a href={"#"} target="_blank"><Share style={{ color: "#282e3c" }}></Share></a>
            //             );
            //         }
            //     }
            // },
            {
                name: "",
                label: "Additional Attachments",
                download: false,
                options: {
                    filter: false,
                    sort: false,
                    customBodyRender: (value, t, f, v) => {
                        let navLink = this.props.folderPath;
                        return (
                            <a href={"#"} onClick={() => { this.setState({ hideDialog: false, selectedDRS: t.rowData[3], fileMetadata: [] }); }}><AttachFile style={{ color: "#282e3c" }}></AttachFile></a>
                        );
                    }
                }
            },
            {
                name: "",
                label: "Approval",
                download: false,
                options: {
                    filter: false,
                    sort: false,
                    customBodyRender: (value, t, f, v) => {
                        let navLink = this.props.folderPath;
                        let status = t.rowData[4];
                        if (status == "Drafted" || status == "Rejected")
                            return (
                                <a title="Send for Approval" href={"#"}
                                    onClick={() => { this.setState({ hideApprovalPane: false, selectedDRS: t.rowData[3], selectedFileID: t.rowData[0] }); }}>
                                    {"Start Approval"}
                                </a>
                            );


                    }
                }
            },
            {
                name: "",
                label: "Approval Status",
                download: false,
                options: {
                    filter: false,
                    sort: false,
                    customBodyRender: (value, t, f, v) => {
                        let navLink = this.props.folderPath;
                        let status = t.rowData[4];
                        if (status == "Drafted")
                            return "To Start";
                        else if (status == "Approved") {
                            return "Approved";

                        }
                        else if (status == "Rejected") {
                            return "Rejected";
                        }
                        else {
                            return "In Progress";
                        }
                    }
                }
            },
            {
                name: "",
                label: "DRS Library",
                download: false,
                options: {
                    filter: false,
                    sort: false,
                    display: false,
                    viewColumns: true,
                    customBodyRender: (value, t, f, v) => {
                        let navLink = this.props.folderPath;
                        return (
                            <a href={navLink} target="_blank"><Folder style={{ color: "#199ad7" }}></Folder></a>
                        );
                    }
                }
            }

        ];
        if (window.location.href.indexOf("Reload") > 0) {
            window.location.href = window.location.pathname;
        }
        this.getListItems(this);

    }
    // UniqueID0  -- 19


    // private handleCopy = ()=>{

    // }

    private handleClick = () => {
        this.setState({
            hideDateRangeModal: false
        });
    }
    public render(): React.ReactElement<IDashboardProps> {
        let classes: any = useStyles;
        let backDropClasses: any = backdropStyles;
        const optionsnew = {
            filterType: 'multiselect',
            selectableRows: false,
            print: false,
            viewColumns: true,
            download: true,
            rowsPerPage: 20,
            sortOrder: {
                name: 'Id',
                direction: 'desc'
            },
            rowsPerPageOptions: [20, 40, 60, 100, 200],
            textLabels: {
                body: {
                    noMatch: this.state.loading ? (<CircularProgress />) : (<div>No Requests Found</div>)
                }
            },
            onFilterChange: (changedColumn, filterList) => {
                console.log(changedColumn, filterList);
            },
            // customToolbar: () => {
            //     return (

            //         <Tooltip title={"Date Range Search"}>
            //             <IconButton onClick={this.handleClick}>
            //                 <DateRangeIcon />
            //             </IconButton>
            //         </Tooltip>

            //     );
            // },
            downloadOptions: { filename: 'ChangeControl.csv', separator: ',' },

            onDownload: (buildHead, buildBody, columns1, data) => {
                if (true) {
                    var csv = `${buildHead(columns1)}${buildBody(data)}`.trim();
                    this.downloadCSV(csv, 'CRRequests.csv');
                }

                return false;
            }

        };
        const { acceptedFiles, searchRequestitems, AllDocuments, selectedDRS, loading, searchStartDate, searchEndDate, dateErrorMsg } = this.state;
        const handleDrop = (acceptedFilesOne) => {
          console.log('Accepted files:', acceptedFilesOne);
        };
        return (
            <>
                <div className={styles.dashboard}>
                    <MUIDataTable
                        data={searchRequestitems}
                        columns={this.gridColumns}
                        options={optionsnew}
                    />
                </div>


                {/* <Dialog
                        hidden={this.state.hideDialog}
                        onDismiss={() => { }}
                        dialogContentProps={dialogContentProps}
                        modalProps={modelProps}
                    > */}

                {/* <div className={styles.row}>

                            <div className={styles.columrowhalf}>
                                <Label>Additional users who will receive notifications</Label>
                                <PeoplePicker
                                    context={this.props.context}
                                    personSelectionLimit={6}
                                    ensureUser={true}
                                    groupName={""} // Leave this blank in case you want to filter from all users
                                    showtooltip={true}
                                    required={true}
                                    onChange={this._getApproverUsers.bind(this)}
                                    showHiddenInUI={false}
                                    principalTypes={[PrincipalType.User]}
                                    resolveDelay={1000}
                                />

                            </div>

                        </div> */}

                {/* <div className={styles.divcontainer}>
                            <div className={styles.divhr}>
                                <h6 className={styles.divheading}>ADDITIONAL ATTACHMENTS (IF NEEDED)</h6>
                            </div>

                        </div> */}

                <Panel
                    isOpen={!this.state.hideDialog}
                    onDismiss={this.candelModel}
                    headerText=''
                    closeButtonAriaLabel="Close"
                    onRenderFooter={() => {
                        return <div className={styles.row}>
                            <div style={{ textAlign: "center" }}>
                                <PrimaryButton onClick={this.saveAttachments.bind(this)} text="Save" style={{ margin: "10px 15px" }} />
                                <DefaultButton onClick={this.candelModel.bind(this)} text="Cancel" style={{ margin: "10px 15px" }} />
                            </div>
                        </div>;
                    }}
                    // Stretch panel content to fill the available height so the footer is positioned
                    // at the bottom of the page
                    isFooterAtBottom={true}
                    type={PanelType.medium}
                >

                    <div className={styles.newRequest}>
                        <Backdrop className={backDropClasses.backdrop} style={{ zIndex: 9 }} open={this.state.loading}>
                            <CircularProgress color="inherit" />
                        </Backdrop>
                        <div className={classes.root}>
                            <div className={styles.divcontainer}>
                                <div className={styles.divhr}>
                                    <h6 className={styles.divheading}>ADDITIONAL ATTACHMENTS (IF NEEDED)</h6>
                                </div>

                            </div>

                            <div className={styles.boxRow}>
                                <div className={styles.row}>

                                    <Dropzone onDrop={this.onDrop.bind(this)}>
                                        {({
                                            getRootProps,
                                            getInputProps,
                                            isDragReject,
                                            isDragActive,
                                            fileRejections
                                        }) => {
                                            // const isFileTooLarge = fileRejections.length > 0 && fileRejections[0].file.size > this.props.FileUploadLimit;
                                            console.log("Is drag active: ", isDragActive);
                                            console.log("Is Drag reject: ", isDragReject)
                                            return (
                                                <div>
                                                    <div className="container">
                                                        <div {...getRootProps({ className: styles.dropzone })}>
                                                            <input {...getInputProps()} />
                                                            {!isDragActive && <p>Click to select files.<Info style={{ position: 'relative', top: '5px' }} fontSize="small"></Info> </p>}
                                                            {isDragActive && !isDragReject && <p>Drop files here</p>}
                                                        </div>
                                                    </div>
                                                    {/* {(acceptedFiles.length > 0 ?
                                                        <aside>
                                                            <br /><br />
                                                            <TableContainer component={Paper}>
                                                                <Table size="small" aria-label="attachment-Table">
                                                                    <TableHead>
                                                                    </TableHead>
                                                                    <TableBody>
                                                                        {this.renderFiles()}
                                                                    </TableBody>
                                                                </Table>
                                                            </TableContainer>
                                                        </aside>
                                                        : "")} */}

                                                    <div className="fileattachment">
                                                        <br /><br />
                                                        <div className={styles.divhr}>
                                                            <h6 className={styles.divheading}>Files Attached :</h6>
                                                        </div>
                                                        <div>

                                                            <TableContainer component={Paper}>
                                                                <Table className={classes.table} size="small" aria-label="a dense table">
                                                                    <TableHead>
                                                                        <TableRow style={{ background: "#b4f7cf" }}>
                                                                            <TableCell>File Name</TableCell>
                                                                            <TableCell>File Type</TableCell>
                                                                        </TableRow>
                                                                    </TableHead>
                                                                    <TableBody>
                                                                        {AllDocuments.filter(_f => _f.DRSNumber == selectedDRS).map((row) => (
                                                                            <TableRow key={row.Title}>
                                                                                <TableCell component="th" scope="row">
                                                                                    <div>
                                                                                        <a href={window.location.origin + row.fileURL} target="_blank">{row.Title}</a>
                                                                                    </div>
                                                                                </TableCell>
                                                                                <TableCell component="th" scope="row">
                                                                                    <div>
                                                                                        {row.Name.substring(row.Name.lastIndexOf(".") + 1, row.Name.length)}
                                                                                    </div>
                                                                                </TableCell>
                                                                            </TableRow>
                                                                        ))}
                                                                    </TableBody>
                                                                </Table>
                                                            </TableContainer>

                                                            {(acceptedFiles.length > 0 ?
                                                                <aside>
                                                                    <div className={styles.divhr}>
                                                                        <h6 className={styles.divheading}>PENDING ATTACHMENTS</h6>
                                                                    </div>
                                                                    <TableContainer component={Paper}>
                                                                        <Table aria-label="attachment-Table">
                                                                            <TableHead>

                                                                            </TableHead>
                                                                            <TableBody>
                                                                                {this.renderFiles()}
                                                                            </TableBody>
                                                                        </Table>
                                                                    </TableContainer>
                                                                </aside>
                                                                : "")}

                                                        </div>
                                                    </div>
                                                </div>
                                            );
                                        }}
                                    </Dropzone>

                                </div>
                            </div>
                        </div>
                    </div>

                </Panel>


                <Panel
                    isOpen={!this.state.hideApprovalPane}
                    onDismiss={this.candelModel}
                    headerText=''
                    closeButtonAriaLabel="Close"
                    onRenderFooter={() => {
                        return <div className={styles.row}>
                            <div style={{ textAlign: "center" }}>
                                <PrimaryButton disabled={this.state.saveApproval} onClick={this.saveApprovers.bind(this)} text="Send for Approval" style={{ margin: "10px 15px" }} />
                                <DefaultButton onClick={this.candelModel.bind(this)} text="Cancel" style={{ margin: "10px 15px" }} />
                            </div>
                        </div>;
                    }}
                    // Stretch panel content to fill the available height so the footer is positioned
                    // at the bottom of the page
                    isFooterAtBottom={true}
                    type={PanelType.medium}

                >
                    <div className={styles.newRequest}>
                        <Backdrop className={backDropClasses.backdrop} style={{ zIndex: 9 }} open={loading}>
                            <CircularProgress color="inherit" />
                        </Backdrop>
                        <div className={classes.root}>
                            <div className={styles.divcontainer}>
                                <div className={styles.divhr}>
                                    <h6 className={styles.divheading}>DRS Approver(s)</h6>
                                </div>

                            </div>
                            <div className={styles.boxRow}>
                                <div className={styles.row}>
                                    <div className={styles.columrowhalf}>
                                        <TextField
                                            id="DRSNumber"
                                            name="DRSNumber"
                                            label="DRS Number"
                                            value={selectedDRS}
                                            readOnly={true}
                                        />
                                    </div>

                                </div>
                                <div className={styles.row}>
                                    <div className={styles.columrowhalf}>
                                        <Label>Recommend Approver</Label>
                                        <PeoplePicker
                                            context={this.props.context}
                                            personSelectionLimit={1}
                                            ensureUser={true}
                                            groupName={""} // Leave this blank in case you want to filter from all users
                                            showtooltip={true}
                                            required={true}
                                            onChange={this._geRecommendUsers.bind(this)}
                                            showHiddenInUI={false}
                                            principalTypes={[PrincipalType.User]}
                                            resolveDelay={1000}
                                        />

                                    </div>

                                </div>
                                <div className={styles.row}>
                                    <div className={styles.columrowhalf}>
                                        <Label>Agree Approver(s)</Label>
                                        <PeoplePicker
                                            context={this.props.context}
                                            personSelectionLimit={6}
                                            ensureUser={true}
                                            groupName={""} // Leave this blank in case you want to filter from all users
                                            showtooltip={true}
                                            required={true}
                                            onChange={this._getAgreeUsers.bind(this)}
                                            showHiddenInUI={false}
                                            principalTypes={[PrincipalType.User]}
                                            resolveDelay={1000}
                                        />

                                    </div>

                                </div>
                                <div className={styles.row}>
                                    <div className={styles.columrowhalf}>
                                        <Label>Perform Approver</Label>
                                        <PeoplePicker
                                            context={this.props.context}
                                            personSelectionLimit={1}
                                            ensureUser={true}
                                            groupName={""} // Leave this blank in case you want to filter from all users
                                            showtooltip={true}
                                            required={true}
                                            onChange={this._getPerformUsers.bind(this)}
                                            showHiddenInUI={false}
                                            principalTypes={[PrincipalType.User]}
                                            resolveDelay={1000}
                                        />

                                    </div>

                                </div>
                                <div className={styles.row}>
                                    <div className={styles.columrowhalf}>
                                        <Label>Input Approver(s)</Label>
                                        <PeoplePicker
                                            context={this.props.context}
                                            personSelectionLimit={6}
                                            ensureUser={true}
                                            groupName={""} // Leave this blank in case you want to filter from all users
                                            showtooltip={true}
                                            required={true}
                                            onChange={this._getInputUsers.bind(this)}
                                            showHiddenInUI={false}
                                            principalTypes={[PrincipalType.User]}
                                            resolveDelay={1000}
                                        />

                                    </div>

                                </div>
                                <div className={styles.row}>
                                    <div className={styles.columrowhalf}>
                                        <Label>Decision Approver</Label>
                                        <PeoplePicker
                                            context={this.props.context}
                                            personSelectionLimit={1}
                                            ensureUser={true}
                                            groupName={""} // Leave this blank in case you want to filter from all users
                                            showtooltip={true}
                                            required={true}
                                            onChange={this._getApproverUsers.bind(this)}
                                            showHiddenInUI={false}
                                            principalTypes={[PrincipalType.User]}
                                            resolveDelay={1000}
                                        />

                                    </div>

                                </div>
                            </div>
                        </div>
                    </div>

                </Panel>


            </>
        );
    }


    private getListItems = async (myProps: any): Promise<void> => {
        let listItems: any = [];
        let fileItems: any = [];
        try {

            // let expand = "Requester,AssignedTo,Author,AttachmentFiles";
            let filter = "";// "ListItemAllFields/DocType eq 'DRS Document'";

            if (this.props.Filter != "" && this.props.Filter != null) {
                //filter = "ListItemAllFields/Project_x0020_Name eq '" + this.props.Filter + "'";
                filter = this.props.Filter;
            }
            listItems = await myProps.props.listServices.GetfileFromLibrary(this.props.folderPath, "", filter);

            if (listItems) {
                if (listItems.length > 0) {
                    listItems.forEach(element => {
                        let reqItem = {}

                        reqItem["fileURL"] = element.ServerRelativeUrl;
                        reqItem["Id"] = element.ListItemAllFields.Id;
                        reqItem["DRSName"] = element.ListItemAllFields.DRSName;
                        reqItem["DRSNumber"] = element.ListItemAllFields.DRSNumber;
                        reqItem["DRSStatus"] = element.ListItemAllFields.DRS_x0020_Status;
                        reqItem["ProjectName"] = element.ListItemAllFields.Project_x0020_Name;
                        reqItem["Project_Number"] = element.ListItemAllFields.Project_x0020_Number;
                        reqItem["DocType"] = element.ListItemAllFields.DocType;
                        reqItem["Title"] = element.ListItemAllFields.Title;
                        reqItem["Name"] = element.Name;

                        if (element.ListItemAllFields.Target_x0020_Date)
                            reqItem["TargetDate"] = element.ListItemAllFields.Target_x0020_Date; //moment(element.ListItemAllFields.Target_x0020_Date).format("DD/MM/YYYY");
                        else
                            reqItem["TargetDate"] = "";

                        fileItems.push(reqItem);
                    });
                }

                let allItems = fileItems;
                let filterItems = fileItems.filter(_i => _i.DocType == "DRS Document");

                this.setState({
                    items: filterItems,
                    loading: false,
                    searchRequestitems: filterItems,
                    AllDocuments: allItems
                });
            }


        } catch (err) {
            console.log("Error getting in List Iteam", err);
        }
        return fileItems;
    }

    private buildCSV(columns, data, options) {
        var replaceDoubleQuoteInString = columnData =>
            typeof columnData === "string"
                ? columnData.replace(/\"/g, '""')
                : columnData;

        var buildHead = columns1 => {
            return (
                columns1
                    .reduce(
                        (soFar, column) =>
                            column.download
                                ? soFar +
                                '"' +
                                this.escapeDangerousCSVCharacters(
                                    replaceDoubleQuoteInString(column.label || column.name)
                                ) +
                                '"' +
                                options.downloadOptions.separator
                                : soFar,
                        ""
                    )
                    .slice(0, -1) + "\r\n"
            );
        };
        var CSVHead = buildHead(columns);

        const buildBody = data1 => {
            return data1
                .reduce(
                    (soFar, row) =>
                        soFar +
                        '"' +
                        row.data
                            .filter((_, index) => columns[index].download)
                            .map(columnData => replaceDoubleQuoteInString(columnData))
                            .join('"' + options.downloadOptions.separator + '"') +
                        '"\r\n',
                    []
                )
                .trim();
        };
        var CSVBody = buildBody(data);

        var csv = options.onDownload
            ? options.onDownload(buildHead, buildBody, columns, data)
            : `${CSVHead}${CSVBody}`.trim();

        return csv;
    }

    private downloadCSV(csv, filename) {
        const blob = new Blob([csv], { type: "text/csv" });

        /* taken from react-csv */
        if (navigator && navigator.msSaveOrOpenBlob) {
            navigator.msSaveOrOpenBlob(blob, filename);
        } else {
            const dataURI = `data:text/csv;charset=utf-8,${csv}`;

            const URL = window.URL;
            const downloadURI =
                typeof URL.createObjectURL === "undefined"
                    ? dataURI
                    : URL.createObjectURL(blob);

            let link = document.createElement("a");
            link.setAttribute("href", downloadURI);
            link.setAttribute("download", filename);
            link.dataset.interception = "off";
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    }

    private createCSVDownload(columns, data, options) {
        const csv = this.buildCSV(columns, data, options);

        if (options.onDownload && csv === false) {
            return;
        }

        this.downloadCSV(csv, options.downloadOptions.filename);
    }

    private escapeDangerousCSVCharacters(data) {
        if (typeof data === "string") {
            // Places single quote before the appearance of dangerous characters if they
            // are the first in the data string.
            return data.replace(/^\+|^\-|^\=|^\@/g, "'$&");
        }

        return data;
    }

    //handle Save Additional Attachments
    private saveAttachments = async () => {
        this.setState({ loading: true });
        //Master DRS Library
        if (this.state.fileMetadata.length > 0) {
            let tempFiles = [];
            this.state.fileMetadata.forEach(async element => {

                let listAttachRes = await this.props.listServices.UploadFiletoLibrary(this.props.folderPath, element);
            });

            //  let listAttachRes = await this.props.listServices.AddListAttachments(this.props.RequestListID, requestRes.data.ID, fileInfos);
        }
        setTimeout(() => {
            this.setState({ hideDialog: true, loading: false, selectedDRS: "", acceptedFiles: [], fileMetadata: [] });
        }, 5000);

    }

    //handle Save Approvers
    private saveApprovers = async () => {
        const { RecommendApprovers, AgreeApprovers, PerformApprovers, InputApprovers, DecisionApprovers, searchRequestitems } = this.state;
        this.setState({ loading: true, saveApproval: true });


        if (RecommendApprovers != null && AgreeApprovers.length > 0 && PerformApprovers != null && DecisionApprovers != null) {


            let filePath = searchRequestitems.filter(_i => _i.Id == this.state.selectedFileID)[0].fileURL;

            //Master DRS Library
            let requestDetail = {};
            requestDetail["Title"] = this.state.selectedDRS;
            requestDetail["FilePath"] = window.location.origin + filePath;
            if (RecommendApprovers != undefined && RecommendApprovers != null)
                requestDetail["Recommend_ApproversId"] = RecommendApprovers;

            if (AgreeApprovers != undefined && AgreeApprovers.length > 0)
                requestDetail["Agree_ApproversId"] = { "results": AgreeApprovers };

            if (PerformApprovers != undefined && PerformApprovers != null)
                requestDetail["Perform_ApproversId"] = PerformApprovers;

            if (InputApprovers != undefined && InputApprovers.length > 0)
                requestDetail["Input_ApproversId"] = { "results": InputApprovers };

            if (DecisionApprovers != undefined && DecisionApprovers != null)
                requestDetail["Decision_ApproversId"] = DecisionApprovers;

            let requestRes = await this.props.listServices.CreateListItem(this.props.DRSApprovalsListId, requestDetail);

            let drsItemUpdate = {};
            drsItemUpdate["DRS_x0020_Status"] = "Submitted for Approval";

            let updateDRS = await this.props.listServices.UpdateListItemByTitle("Master DRS Library", drsItemUpdate, this.state.selectedFileID);

            setTimeout(() => {
                this.setState({
                    hideApprovalPane: true, loading: false,
                    RecommendApprovers: [], AgreeApprovers: [], PerformApprovers: [], InputApprovers: [], DecisionApprovers: []
                });
                window.location.reload();
            }, 5000);
        }
        else {
            alert("Please select all Approvers for DRS Approval");
        }

    }


    private candelModel = async () => {
        this.setState({
            hideDialog: true, hideApprovalPane: true,
            selectedDRS: "", acceptedFiles: [], fileMetadata: [],
            RecommendApprovers: [], AgreeApprovers: [], PerformApprovers: [], InputApprovers: [], DecisionApprovers: []
        });
    }

    private _geRecommendUsers = (items: any[]) => {

        let additionalUsers = [];
        if (items.length > 0) {
            items.map(user => {
                additionalUsers.push(user.id);
            });
            this.setState({ RecommendApprovers: items[0].id });
        }
        else {
            additionalUsers = [];
            this.setState({ RecommendApprovers: null });
        }

    }

    private _getAgreeUsers = (items: any[]) => {

        let additionalUsers = [];
        if (items.length > 0) {
            items.map(user => {
                additionalUsers.push(user.id);
            });
        }
        else {
            additionalUsers = [];
        }

        this.setState({ AgreeApprovers: additionalUsers });

    }
    private _getPerformUsers = (items: any[]) => {
        let additionalUsers = [];
        if (items.length > 0) {
            items.map(user => {
                additionalUsers.push(user.id);
            });
            this.setState({ PerformApprovers: items[0].id });
        }
        else {
            additionalUsers = [];
            this.setState({ PerformApprovers: null });
        }

        // this.setState({ PerformApprovers: additionalUsers });

    }
    private _getInputUsers = (items: any[]) => {

        let additionalUsers = [];
        if (items.length > 0) {
            items.map(user => {
                additionalUsers.push(user.id);
            });
        }
        else {
            additionalUsers = [];
        }

        this.setState({ InputApprovers: additionalUsers });

    }
    private _getApproverUsers = (items: any[]) => {

        let additionalUsers = [];
        if (items.length > 0) {
            items.map(user => {
                additionalUsers.push(user.id);
            });
            this.setState({ DecisionApprovers: items[0].id });
        }
        else {
            additionalUsers = [];
            this.setState({ DecisionApprovers: null });
        }

        // this.setState({ DecisionApprovers: additionalUsers });

    }


    private onDrop = (selectedFiles, rejectedFiles) => {

        var fileInfos: any[] = [];
        let filteredSelectedFiles = [];
        let fileData = [];


        if (this.state.acceptedFiles.length > 0) {
            selectedFiles.forEach(element => {
                let isFileExists = this.state.acceptedFiles.filter(val => val.Name == element.name);
                if (isFileExists.length == 0) {
                    filteredSelectedFiles.push(element);
                }
                else {
                    let msg = "Sorry, " + element.name + " file already exists with same name in the New Attachments";
                    // NotificationManager.warning(msg, 'Warning', 3000);
                    alert(msg);
                }
            });
        }
        else {
            filteredSelectedFiles = selectedFiles;
        }

        if (filteredSelectedFiles.length > 0) {
            //this.setState({ fileUploading: true });
            filteredSelectedFiles.map(file => {


                const reader = new FileReader();

                reader.onabort = () => console.log('file reading was aborted');
                reader.onerror = () => console.log('file reading has failed');
                reader.onload = () => {
                    let pattern = /[/\\~#%&*{}/:<>?|\"-]/g;
                    let replacement = "_";
                    let regex = new RegExp(pattern);

                    let fileName = file.name.replace(regex, replacement).replace(/_+/g, '_');
                    fileInfos.push({
                        "Name": fileName,
                        "Content": reader.result,
                        "FileName": ""
                    });

                    let _metaData = {
                        Title: fileName,
                        DRSNumber: this.state.selectedDRS,
                        DocType: "Additional Document"
                    };

                    let _fileInfor = {
                        name: fileName,
                        content: file
                    };

                    fileData.push({
                        file: _fileInfor,
                        metaData: _metaData
                    });


                    // Do whatever you want with the file contents
                    const binaryStr = reader.result;


                    if (filteredSelectedFiles.length == fileInfos.length) {
                        this.setState({
                            acceptedFiles: [...this.state.acceptedFiles, ...fileInfos]//, fileUploading: false
                            , fileMetadata: [...this.state.fileMetadata, ...fileData]
                        });

                    }
                };


                reader.readAsArrayBuffer(file);


            });
        }


    }


    private renderFiles() {

        let files = this.state.acceptedFiles;
        if (files && files.length > 0) {

            return files.map((file, index) => {
                return (
                    <TableRow key={index}>
                        <TableCell className={styles.column}>{index + 1}</TableCell>
                        <TableCell className={styles.column}>{file.Name}</TableCell>
                        <TableCell className={styles.column}><span
                            onClick={() => this.handleIconclickDelete(file.Name)}
                            style={{ cursor: "pointer" }}
                            placeholder="Click to delete the file">
                            <Tooltip title="Delete attachment"><Delete htmlColor="#ce1b0e" fontSize="small" /></Tooltip>
                        </span></TableCell>
                    </TableRow>
                );
            });
        }
    }

    private handleIconclickDelete = async (fileName: any) => {
        try {
            let rowsArray = this.state.acceptedFiles;
            var rows = this.state.acceptedFiles.filter(item => item.Name !== fileName);
            this.setState({ acceptedFiles: rows });

        } catch (error) {
            console.log("Error in React Table handle Remove Row : " + error);
        }

    }





}

const dropzoneStyles: React.CSSProperties = {
  border: '2px dashed #888',
  borderRadius: '5px',
  padding: '40px',
  textAlign: 'center', // Set a valid value of type TextAlignProperty, like "center"
};
