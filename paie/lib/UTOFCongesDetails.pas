unit UTOFCongesDetails;

interface
uses Classes,SysUtils,
     Utob,Utof,HCtrls,HMsgBox,HEnt1

{$IFNDEF EAGLCLIENT}
      ,{$IFNDEF DBXPRESS} dbTables {$ELSE} uDbxDataSet {$ENDIF} //HDB,DBCtrls,Fiche,fichlist,Mul,HQry,Fe_Main,
{$ELSE}
      // eFiche,efichlist,eMul,MaineAgl,
{$ENDIF}

     ;



Type
     TOF_PGCongesDetails = Class (TOF)
       procedure OnArgument(stArgument: String); override;
       procedure OnLoad ;                        override ;
       procedure OnUpdate ;                              override ;

       Private
       GblSalarie : String;
     End;

implementation
Uses PgCongesPayes;

{ TOF_PGCongesDetails }

procedure TOF_PGCongesDetails.OnArgument(stArgument: String);
Var
DateArrete : TDateTime;
Q          : TQuery ;
begin
  inherited;
  GblSalarie := ReadTokenSt(StArgument);
  SetControlText('SALARIE',GblSalarie);
  Q:=OpenSQL('SELECT MAX(PCN_DATEVALIDITE) DATEVALIDITE FROM ABSENCESALARIE '+
           'WHERE PCN_TYPEMVT="CPA" AND PCN_SALARIE="'+GblSalarie+'" '+
//           'AND (PCN_TYPECONGE="ACQ" OR PCN_TYPECONGE="ACS" OR PCN_TYPECONGE="ACA" OR PCN_TYPECONGE="REL"  '
           'AND PCN_PERIODECP<2',True);
  if Not Q.Eof then DateArrete := Q.FindField('DATEVALIDITE').AsDateTime
  else              DateArrete := IDate1900;
  Ferme(Q);
  SetControlText('DATEARRETE',DateToStr(DateArrete));

end;

procedure TOF_PGCongesDetails.OnLoad;
Var
  DateArrete,DateDN1,DateFN1,DateDN : TDateTime;
  Tob_Prov,T                        : Tob;
begin
  inherited;
  if IsValidDate(GetControlText('DATEARRETE')) then
    DateArrete := StrToDate(GetControlText('DATEARRETE'))
  else
     Begin
     PGiBox('Le format de date la date d''arrêté est incorrect.',Ecran.Caption);
     Exit;
     End;

  Tob_Prov:=CalculeProvisionCp('AND PSA_SALARIE="'+GblSalarie+'"',DateArrete,DateDN1,DateFN1,DateDN);
  FreeAndNil(T);
  if Assigned(Tob_Prov) then T:= Tob_Prov.FindFirst(['SALARIE'],[GblSalarie],False);
  if Assigned(T) then
    Begin
    SetControlProperty('GBRELN1','Caption','Reliquat N-1 au '+DateToStr(DateDN1));
    SetControlProperty('GBDETAILN1','Caption','Période N-1 au '+DateToStr(DateDN1));
    SetControlProperty('GBDETAILN' ,'Caption','Période N au '+DateToStr(DateDN) );

    SetControlText('CPBASERELN1'    ,T.GetValue('CPBASERELN1'));
    SetControlText('CPMOISRELN1'    ,T.GetValue('CPMOISRELN1'));
    SetControlText('CPACQUISRELN1'  ,T.GetValue('CPACQUISRELN1'));
    SetControlText('CPPRISRELN1'    ,T.GetValue('CPPRISRELN1'));
    SetControlText('CPRESTRELN1'    ,T.GetValue('CPRESTRELN1'));
    SetControlText('CPPROVRELN1'    ,T.GetValue('CPPROVRELN1'));

    SetControlText('CPBASEN1'       ,T.GetValue('CPBASEN1'));
    SetControlText('CPMOISN1'       ,T.GetValue('CPMOISN1'));
    SetControlText('CPACQUISN1'     ,T.GetValue('CPACQUISN1'));
    SetControlText('CPPRISN1'       ,T.GetValue('CPPRISN1'));
    SetControlText('CPRESTN1'       ,T.GetValue('CPRESTN1'));
    SetControlText('CPPROVN1'       ,T.GetValue('CPPROVN1'));

    SetControlText('CPBASEN'        ,T.GetValue('CPBASEN'));
    SetControlText('CPMOISN'        ,T.GetValue('CPMOISN'));
    SetControlText('CPACQUISN'      ,T.GetValue('CPACQUISN'));
    SetControlText('CPPRISN'        ,T.GetValue('CPPRISN'));
    SetControlText('CPRESTN'        ,T.GetValue('CPRESTN'));
    SetControlText('CPPROVN'        ,T.GetValue('CPPROVN'));

    SetControlText('CPPROVISION'    ,T.GetValue('CPPROVISION'));
    End;

end;

procedure TOF_PGCongesDetails.OnUpdate;
begin
  inherited;
OnLoad;
end;

Initialization
registerclasses([TOF_PGCongesDetails]) ;
end.
 
