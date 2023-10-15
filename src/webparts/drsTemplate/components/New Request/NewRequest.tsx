import * as React from 'react';
import { INewRequestProps, ICarrierOptions } from "./INewRequestProps";
import { Button, MenuItem, makeStyles, createStyles, Theme } from "@material-ui/core";
import styles from "./NewRequest.module.scss";
import { NotificationContainer } from 'react-notifications';
import 'react-notifications/lib/notifications.css';
import { NotificationManager } from 'react-notifications';
import { Delete, Info } from '@material-ui/icons';
import { IAttachmentFileInfo } from '@pnp/sp/attachments';
import { CircularProgress, Checkbox } from '@material-ui/core';
import { TableContainer, Table, TableBody, TableCell, TableHead, TableRow, Paper, Tooltip, Backdrop } from '@material-ui/core';
//import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { HttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http';

import Select from 'react-select';
import * as moment from 'moment';

import { DatePicker, Dropdown, IDropdownOption, Label, TeachingBubbleBase, TextField } from 'office-ui-fabric-react';


interface FilmOptionType {
    title: string;
    inputValue?: string;

}

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
type Order = 'asc' | 'desc';

var uploadedDocsList = [];
export class NewRequest extends React.Component<INewRequestProps, any> {

    public ProjectOptions = [];
    public ProjectNumberOptions = [];
    constructor(props: INewRequestProps) {
        super(props);

        this.state = {
            NewItem: {},
            formloading: true,
            loading: false,
            acceptedFiles: [],
            FilteredListItems: [],
            saveLoading: false, isSaveDisabled: true, saveValidating: false,
            projectTitle: [], projectNumbers: [], projectOptions: [],
            isDataSelected: false,
            order: 'asc', orderBy: null

        };

        this.getProjectTitles = this.getProjectTitles.bind(this);
        this.getProjectNumbers = this.getProjectNumbers.bind(this);

    }

    public componentDidMount = async (): Promise<void> => {
        console.log("Start");
        //   await this.getProjectNumbers();
        await this.getProjectTitles();


    }


    private onDrop = (selectedFiles, rejectedFiles) => {

        var fileInfos: any[] = [];
        let filteredSelectedFiles = [];

        if (this.state.acceptedFiles.length > 0) {
            selectedFiles.forEach(element => {
                let isFileExists = this.state.acceptedFiles.filter(val => val.Name == element.name);
                if (isFileExists.length == 0) {
                    filteredSelectedFiles.push(element);
                }
                else {
                    let msg = "Sorry, " + element.name + " file already exists with same name in the New Attachments";
                    NotificationManager.warning(msg, 'Warning', 3000);
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

                    // Do whatever you want with the file contents
                    const binaryStr = reader.result;


                    if (filteredSelectedFiles.length == fileInfos.length) {
                        this.setState({
                            acceptedFiles: [...this.state.acceptedFiles, ...fileInfos]//, fileUploading: false
                        });

                    }
                };


                reader.readAsArrayBuffer(file);


            });
        }


    }



    private OnDocsDelete = (index: number) => {
        uploadedDocsList.splice(index, 1);
        this.setState({
            uploadDocInfo: uploadedDocsList,
        });
    }


    private handleBack = (): void => {
        // window.location.hash = '#/requests/my';
        //window.history.back();
        if (window.history.go(-1) == undefined)
            window.location.hash = '#/';
        else
            window.history.go(-1);
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
    private async handleSave() {
        try {
            const { NewItem } = this.state;
            if (NewItem.DRSName != undefined && NewItem.ProjectName != undefined) {
                this.setState({ loading: true });
                let flowURL = this.props.FlowURL;//"https://prod-23.australiasoutheast.logic.azure.com:443/workflows/9bda11b287c94bb1a3a7539b2db81f13/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=M8VlfwlRoCEWbTEm8-bZQROoM6v6eDYFyyVruN_n06M";
                let reqItem = {};
                reqItem["DRSName"] = this.state.NewItem.DRSName;
                reqItem["Project Name"] = this.state.NewItem.ProjectName;
                reqItem["Project Nember"] = this.state.NewItem.ProjectNumber;
                reqItem["Phase"] = this.state.NewItem.Phase;
                reqItem["TargetDate"] = moment(new Date()).toISOString();
                reqItem["DocumentNo"] = this.state.NewItem.DocumentNo;

                const requestHeaders: Headers = new Headers();
                requestHeaders.append('Content-type', 'application/json');
                requestHeaders.append('Cache-Control', 'no-cache');

                const httpClientOptions: IHttpClientOptions = {
                    body: JSON.stringify(reqItem),
                    headers: requestHeaders
                };


                await this.props.context.httpClient.post(
                    flowURL, HttpClient.configurations.v1, httpClientOptions
                ).then((response: HttpClientResponse) => {
                    console.log("REST API response received.");
                    setTimeout(() => {
                        this.setState({ loading: false });
                        window.location.hash = '#/';
                    }, 10000);

                    //return response.json();
                });
            }
            else {
                alert("Please fill testing all required fields to Submit the request");
            }

        }
        catch (err) {
            console.log(err);
            //  NotificationManager.error('Error Submiting the request', 'Error!', 5000);
            setTimeout(() => {
                this.setState({ loading: false });
                window.location.hash = '#/';
            }, 10000);
        }
    }

    private getProjectTitles = async () => {
        let _projects = [];

        _projects = await this.props.listServices.GetItemsByListId(this.props.ProjectListId, "Title,ProjectName", "", "", "", true);

        if (_projects.length > 0) {
            _projects.map(item => {
                // let option = { key: item.ProjectName, text: item.ProjectName };
                let option = { value: item.ProjectName, label: item.ProjectName };
                this.ProjectOptions.push(option);
            });
            this.setState({ projectTitle: _projects, projectOptions: this.ProjectOptions });

        }
    }

    private getProjectNumbers = async () => {
        let _projectsNumbers = [];
        _projectsNumbers = await this.props.listServices.GetItemsByListId(this.props.ProjectNumberListId, "Title", "", "", "", true);

        if (_projectsNumbers.length > 0) {
            _projectsNumbers.map(item => {
                let option = { key: item.Title, text: item.Title };
                this.ProjectNumberOptions.push(option);
            });
            this.setState({ projectNumbers: _projectsNumbers });

        }
    }


    private _onSelectDate = (date: Date | null | undefined): void => {

        const NewItem = { ...this.state.NewItem, ["TargetDate"]: date };
        this.setState(() => ({ NewItem }));

    }

    private _onFormatDate = (date: Date): string => {
        return date.getDate() + '/' + (date.getMonth() + 1) + '/' + date.getFullYear();
    }

    public render(): React.ReactElement<INewRequestProps> {
        let classes: any = useStyles;
        let backDropClasses: any = backdropStyles;
        const handleChange = (event: React.ChangeEvent<{ name?: string; value: unknown }>) => {
            const NewItem = { ...this.state.NewItem, [event.target.name]: event.target.value };
            this.setState(() => ({ NewItem }));
        };
        const handleDropDownChanges = (filedName: string, item: IDropdownOption): void => {
            const NewItem = { ...this.state.NewItem, [filedName]: item.text };
            this.setState(() => ({ NewItem }));
        };
        const handleTextChange = (event: React.ChangeEvent<HTMLInputElement>) => {
            const NewItem = { ...this.state.NewItem, [event.target.name]: event.target.value };
            this.setState(() => ({ NewItem }));
        };

        const handleSelectProject = (item: any) => {

            let selectedProject = this.state.projectTitle.filter(_i => _i.ProjectName == item.value);
            const NewItem = { ...this.state.NewItem, ["ProjectName"]: item.value, ["ProjectNumber"]: selectedProject[0].Title, ["ProjectOption"]: item };
            this.setState(() => ({ NewItem }));
        }



        //  const [APvalue, setAPValue] = React.useState(null);
        const { NewItem, formloading, acceptedFiles, isSaveDisabled, saveLoading, saveValidating } = this.state;

        return (
            <div className={styles.newRequest}>
                <Backdrop className={backDropClasses.backdrop} style={{ zIndex: 9 }} open={this.state.loading}>
                    <CircularProgress color="inherit" />
                </Backdrop>
                <div className={classes.root}>
                    <div className={styles.divcontainer}>
                        <div className={styles.divhr}>
                            <h6 className={styles.divheading}>Decision Record Details</h6>
                        </div>
                    </div>
                    <div className={styles.boxRow}>
                        <div className={styles.row}>
                            <div className={styles.columrowhalfInvoice}>
                                <TextField
                                    id="DRSName"
                                    name="DRSName"
                                    label="DRS Name"
                                    onChange={handleTextChange}
                                    required
                                    description={"DRS Owner to enter the DRS name"}

                                />
                            </div>

                            <div className={styles.columrowhalfInvoice}>
                                <Label title='Select correct project name from the dropdown list provided' required>Project Name</Label>
                                {/* <Dropdown
                                title='Select correct project name from the dropdown list provided'
                                    placeholder="Select an Project Name"
                                    label="Project Name"
                                    options={this.ProjectOptions}
                                    required={true}
                                    onChange={(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption) => handleSelectProject(item)}

                                /> */}

                                <Select
                                    title='Select correct project name from the dropdown list provided'
                                    classNamePrefix="Select"
                                    isClearable={true}
                                    isSearchable={true}
                                    value={NewItem.ProjectOption}
                                    options={this.state.projectOptions}
                                    onChange={handleSelectProject.bind(this)}

                                />
                            </div>
                        </div>

                        <div className={styles.row}>
                            <div className={styles.columrowhalfInvoice}>

                                <Dropdown
                                    title='Select 1 of the 5 project phases from the dropdown list provided'
                                    placeholder="Select an Phase"
                                    label="Phase"
                                    options={[{ key: 'Assess', text: 'Assess' }, { key: 'Select', text: 'Select' }, { key: 'Define', text: 'Define' }, { key: 'Execute', text: 'Execute' }, { key: 'Operate', text: 'Operate' }]}
                                    onChange={(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption) => handleDropDownChanges("Phase", item)}
                                />

                            </div>
                            <div className={styles.columrowhalfInvoice}>
                                <TextField
                                    id="ProjectNumber"
                                    name="ProjectNumber"
                                    label="Project Number"
                                    onChange={handleTextChange}
                                    value={NewItem.ProjectNumber}
                                    disabled={true}
                                />
                                {/* <Dropdown
                                    placeholder="Select an Project Number"
                                    label="Project Number"
                                    options={this.ProjectNumberOptions}
                                    onChange={(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption) => handleDropDownChanges("ProjectNumber", item)}
                                /> */}
                            </div>

                            <div className={styles.row}>
                                <div className={styles.columrowhalfInvoice}>
                                    <TextField
                                        id="DocumentNo"
                                        name="DocumentNo"
                                        label="Document Number"
                                        onChange={handleTextChange}
                                        required={false}
                                        description={"This should be a Boardwalk/TeamBinder Document Number"}
                                    />
                                </div>
                                <div className={styles.columrowhalfInvoice}>

                                    <TextField
                                        id="TargetDate"
                                        name="TargetDate"
                                        label="Date Created"
                                        disabled={true}
                                        value={moment(new Date()).format("DD/MM/YYYY")}
                                    />

                                </div>
                            </div>


                        </div>

                    </div>
                    {/* <div className={styles.divcontainer}>
                        <div className={styles.divhr}>
                            <h6 className={styles.divheading}>ATTACHMENTS (IF NEEDED)</h6>
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

                                    return (
                                        <div>
                                            <div className="container">
                                                <div {...getRootProps({ className: styles.dropzone })}>
                                                    <input {...getInputProps()} />
                                                    {!isDragActive && <p>Drag 'n' drop some files here, or click to select files.<br></br><Info style={{ position: 'relative', top: '5px' }} fontSize="small"></Info> </p>}
                                                    {isDragActive && !isDragReject && <p>Drop files here</p>}
                                                </div>
                                            </div>
                                            {(acceptedFiles.length > 0 ?
                                                <aside>
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
                                                : "")}

                                        </div>
                                    );
                                }}
                            </Dropzone>

                        </div>
                    </div> */}


                    <div className={styles.row}>
                        <div className={styles.buttonDiv}>
                            <Button color="default" size="large" variant="contained" onClick={this.handleBack.bind(this)} className={classes.button} style={{ margin: '50px', maxWidth: '180px', maxHeight: '40px', minWidth: '180px', minHeight: '40px' }} >Cancel</Button>
                            <span className={styles.wrapper}>
                                <Button
                                    variant="contained"
                                    color="primary"
                                    onClick={this.handleSave.bind(this)}
                                    className={classes.button}
                                    size="large"
                                    style={{ margin: '50px', maxWidth: '180px', maxHeight: '40px', minWidth: '180px', minHeight: '40px' }}
                                >
                                    {saveLoading && (
                                        <CircularProgress size={14}></CircularProgress>
                                    )}
                                    {saveLoading && saveValidating && <div style={{ marginLeft: '5px' }}> Validating...</div>}
                                    {saveLoading && !saveValidating && <div style={{ marginLeft: '5px' }}> Saving...</div>}
                                    {!saveLoading && <div style={{ marginLeft: '5px' }}> Submit</div>}
                                </Button>
                            </span>
                        </div>
                    </div>
                </div >
                <NotificationContainer />

            </div >
        );

    }
}

