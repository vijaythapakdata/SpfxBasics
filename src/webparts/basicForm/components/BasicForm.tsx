import * as React from 'react';
// import styles from './BasicForm.module.scss';
import type { IBasicFormProps } from './IBasicFormProps';
import { IBasicFormState } from './IBasicFormState';
import {Web} from '@pnp/sp/presets/all';
import { PrimaryButton, Slider, TextField } from '@fluentui/react';
import {PeoplePicker, PrincipalType} from '@pnp/spfx-controls-react/lib/PeoplePicker';
export default class BasicForm extends React.Component<IBasicFormProps, IBasicFormState> {
  constructor(props:any){
    super(props);
    this.state={
      Name:'',
      Email:'',
      Age:'',
      Score:'',
      Address:'',
      Manager:[],
      ManagerId:[],
      Admin:'',
      AdminId:''
    }
  }
  private async addItems(){
let web=Web(this.props.siteurl);
await web.lists.getByTitle(this.props.ListName).items.add({
  Title:this.state.Name,
  EmailAddress:this.state.Email,
  Age:parseInt(this.state.Age),
  Score:parseInt(this.state.Score),
  Address:this.state.Address,
  ManagerId:{results:this.state.ManagerId},
  AdminId:this.state.AdminId

}).then((response)=>{
  console.log(response);
  alert('Item added successfully');
  this.setState({
    Name:'',
    Email:'',
    Age:'',
    Score:'',
    Address:'',
    Manager:[],
      ManagerId:[],
      Admin:'',
      AdminId:''
  })

}).catch((err)=>{
  console.log(err);
  alert('Error while adding item');
})
  }
  //Event handling
  private handlChange=(fieldValue:keyof IBasicFormState,value:string|boolean|number):void=>{
this.setState({[fieldValue]:value}as unknown as Pick<IBasicFormState,keyof IBasicFormState>)
  }
  public render(): React.ReactElement<IBasicFormProps> {
    

    return (
    <>
    <TextField label='Name' value={this.state.Name}
    onChange={(_,e)=>this.handlChange('Name',e||'')} iconProps={{iconName:'User'}}
    />
    <TextField label='Age' value={this.state.Age}
    onChange={(_,e)=>this.handlChange('Age',e||'')} 
    />
    <TextField label='Email Address' value={this.state.Email}
    onChange={(_,e:any)=>this.handlChange('Email',e)} iconProps={{iconName:'mail'}}
    />
    <TextField label='Permananet Address' value={this.state.Address}
    onChange={(_,e:any)=>this.handlChange('Address',e)} iconProps={{iconName:'home'}}
    multiline rows={5}
    />
    <Slider label='Score' min={1} max={100} step={1} value={this.state.Score}
    onChange={(value:number)=>this.handlChange('Score',value)}
    />
    <PeoplePicker
    context={this.props.context as any}
    titleText='Manager'
    personSelectionLimit={3}
    principalTypes={[PrincipalType.User]}
    defaultSelectedUsers={this.state.Manager}
    onChange={this._getManagers}
    showtooltip={true}
    resolveDelay={1000}
    webAbsoluteUrl={this.props.siteurl}
    ensureUser={true}
    />
    <PeoplePicker
    context={this.props.context as any}
    titleText='Admin'
    personSelectionLimit={1}
    principalTypes={[PrincipalType.User]}
    defaultSelectedUsers={[this.state.Admin?this.state.Admin:'']}
    onChange={this._getAdmin}
    showtooltip={true}
    resolveDelay={1000}
    webAbsoluteUrl={this.props.siteurl}
    ensureUser={true}
    />
    <br/>
    <PrimaryButton text='Submit' onClick={()=>this.addItems()}
      iconProps={{iconName:'save'}}
      />
    </>
    );
  }
  //Get Manager
  private _getManagers=(item:any):void=>{
    const managers=item.map((items:any)=>items.text)
    const managersId=item.map((items:any)=>items.id)
    this.setState({
      Manager:managers,
      ManagerId:managersId
    });
  }
  //Get Admin
  private _getAdmin=(item:any):void=>{
    if(item.length>0){
      this.setState({
        Admin:item[0].text,
        AdminId:item[0].id
      });
    }else{
      this.setState({
        Admin:'',
        AdminId:''
      });
    }
  }
}
