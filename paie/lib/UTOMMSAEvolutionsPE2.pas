{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 24/03/2004
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : MSAEVOLUTIONSPE2
Mots clefs ... : TOM;MSAEVOLUTIONSPE2
*****************************************************************}
{
PT1 30/11/2006 JL V750 : mdofi valeurs poly employeurs
PT2 01/12/2006 JL V_750 : Ajout prud'hommes
PT3 11/05/2007 JL V_72 FQ 13888 matricule salarié complété automatiquement
PT4 13/11/2007 FC V_80 FQ 14892 Evolutions cahier des charges Octobre 2007
}
Unit UTOMMSAEvolutionsPE2 ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fiche,FichList,
{$else}
     eFiche,eFichList,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOM,UTob,EntPaie,PGOutils2 ;

Type
  TOM_MSAEVOLUTIONSPE2 = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    private
    procedure ChangeTypeEvolution (Sender : TObject);
    procedure ExitEdit(Sender: TObject);   //PT3
    procedure cbComboChange(Sender: TObject); //PT4
    procedure cbValChange(Sender: TObject); //PT4
    end ;

Implementation

procedure TOM_MSAEVOLUTIONSPE2.OnNewRecord ;
begin
  Inherited ;
        SetField('PE2_DATEEFFET',Date);
  //DEB PT4
  SetControlText('PERIODICITE','');
  SetControlText('DUREECTR','');
  //FIN PT4
end ;

procedure TOM_MSAEVOLUTIONSPE2.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_MSAEVOLUTIONSPE2.OnUpdateRecord ;
var MessError : String;
begin
  Inherited ;
        MessError := '';
        If GetField('PE2_SALARIE') = '' then MessError := MessError+'#13#10 - Le matricule du salarié';
        If GetField('PE2_ETABLISSEMENT') = '' then MessError := MessError+'#13#10 - L''établissement du salarié';
        If GetField('PE2_MSAACTIVITE') = '' then MessError := MessError+'#13#10 - L''activité du salarié';
        If GetField('PE2_DATEEFFET') = IDate1900 then MessError := MessError+'#13#10 - La date d''éffet';
        If MessError <> '' then
        begin
                LastError := 1;
                PGIBox('La (ou les) information(s) suivante(s) sont obligatoire(s) :'+MessError,Ecran.Caption);
                exit;
        end;
        SetField('PE2_MSATPSPART',getControltext('SECTIONPRUD')); //PT2
        SetField('PE2_MSACONTRAT',getControltext('COLLEGEPRUD'));

  //DEB PT4
  If GetControlText('PE2_TYPEEVOLMSA') = 'PE23' then
  begin
    SetField('PE2_PERSTECH',getControltext('PERIODICITE'));
    if getControltext('DUREECTR') <> '' then
      SetField('PE2_CODEQUALITE',FloatToStr(StrToFloat(getControltext('DUREECTR')) * 100))
    else
      SetField('PE2_CODEQUALITE',0);
  end;
  //FIN PT4
end ;

procedure TOM_MSAEVOLUTIONSPE2.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_MSAEVOLUTIONSPE2.OnLoadRecord ;
var PContrat,PCompl : TTabSheet;
    Q : TQuery;
begin
  Inherited ;
        PContrat := TTabSheet(GetControl('PE23'));
        PCompl := TTabSheet(GetControl('PE24'));
        PContrat.TabVisible := False;
        PCompl.TabVisible := False;
        SetControlVisible('GRBXPE21',False);
        SetControlVisible('GRBX22',False);
        If GetField('PE2_TYPEEVOLMSA') = 'PE21' then SetControlVisible('GRBXPE21',True);
        If GetField('PE2_TYPEEVOLMSA') = 'PE22' then SetControlVisible('GRBX22',True);
        If GetField('PE2_TYPEEVOLMSA') = 'PE23' then PContrat.TabVisible := True;
        If GetField('PE2_TYPEEVOLMSA') = 'PE24' then PCompl.TabVisible := True;
        Q := OpenSQL('Select O_VILLE from CODEPOST WHERE O_CODEPOSTAL="'+GetField('PE2_CPLIEUTRAV')+'"',True);
        If Not Q.Eof then SetControlCaption('LIBVILLEMSA',Q.FindField('O_VILLE').AsString)
        Else SetControlCaption('LIBVILLEMSA','');
        SetControlText('SECTIONPRUD',GetField('PE2_MSATPSPART'));   //PT2
        SetControlText('COLLEGEPRUD',GetField('PE2_MSACONTRAT'));
        If getField('PE2_DATEEFFET') < EncodeDate(2006, 10, 1) then
          PGIBox('Attention, vous utilisez le cahier des charges applicable à partir #13#10'+
          'du dernier trimestre 2006, certaines informations peuvent donc être absentes',Ecran.Caption);
          
  //DEB PT4
  If GetField('PE2_TYPEEVOLMSA') = 'PE23' then
  begin
    SetControlText('PERIODICITE',GetField('PE2_PERSTECH'));
    SetControlText('DUREECTR',FloatToStr(StrToFloat(GetField('PE2_CODEQUALITE')) / 100));
  end;
  //FIN PT4
end ;

procedure TOM_MSAEVOLUTIONSPE2.OnChangeField ( F: TField ) ;
var Q : TQuery;
begin
  Inherited ;
        If (F.FieldName = 'PE2_SALARIE') and ((DS.State in [dsInsert])) then
        begin
           If GetField('PE2_SALARIE') <> '' Then ExitEdit(THEdit(GetControl('PE2_SALARIE')));//PT3
                If GetField('PE2_SALARIE') <> '' then
                begin
                        Q := OpenSQL('SELECT PSA_SALARIE,PSA_LIBELLE,PSA_NUMEROSS,PSA_ETABLISSEMENT,PSA_PRENOM,PSA_DATENAISSANCE,PSE_MSAACTIVITE FROM SALARIES'+
                        ' LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE'+
                        ' WHERE PSA_SALARIE="'+GetField('PE2_SALARIE')+'"',True);
                        If Not Q.Eof then
                        begin
                                SetField('PE2_ETABLISSEMENT',Q.FindField('PSA_ETABLISSEMENT').AsString);
                                SetField('PE2_NOM',Q.FindField('PSA_LIBELLE').AsString);
                                SetField('PE2_NUMEROSS',Q.FindField('PSA_NUMEROSS').AsString);
                                SetField('PE2_PRENOM',Q.FindField('PSA_PRENOM').AsString);
                                SetField('PE2_DATENAISSANCE',Q.FindField('PSA_DATENAISSANCE').AsDateTime);
                                SetField('PE2_MSAACTIVITE',Q.FindField('PSE_MSAACTIVITE').AsString);
                        end;
                        Ferme(Q);
                end;
        end;
        If F.FieldName = 'PE2_MSAPOLYEMP' then
        begin
                If GetField('PE2_MSAPOLYEMP') = '' then exit;
                If (GetField('PE2_MSAPOLYEMP') <> '0') and (GetField('PE2_MSAPOLYEMP') <> '2') and (GetField('PE2_MSAPOLYEMP') <> '3') then //PT1
                begin
                        PGIBox('Le code poly-employeur ne peut prendre que les valeurs suivantes :'+
                        '#13#10 0 : Mono-employeur ou moins de 2 employeurs agricoles.'+
                        '#13#10 2 : Salarié travaillant pour plusieurs employeurs dont un seul cotise à la CRCCA '+
                        '#13#10     pour la retraite complémentaire des cadres ou assimilés.'+
                        '#13#10 3 : Salarié travaillant pour plusieurs employeurs dont au moins deux cotisent à la CRCCA '+
                        '#13#10     pour la retraite complémentaire des cadres ou assimilés.',Ecran.Caption);
                        SetField('PE2_MSAPOLYEMP','');
                        SetFocusControl('PE2_MSAPOLYEMP');
                end;
        end;
end ;

procedure TOM_MSAEVOLUTIONSPE2.OnArgument ( S: String ) ;
var  {$IFNDEF EAGLCLIENT}
     TypeEvolution : THDBValComboBox;
     {$ELSE}
     TypeEvolution : THValComboBox;
     {$ENDIF}
  Combo : THValComboBox;
  Edit : THEdit;
begin
  Inherited ;
        {$IFNDEF EAGLCLIENT}
        TypeEvolution := THDBValComboBox(GetControl('PE2_TYPEEVOLMSA'));
        {$ELSE}
        TypeEvolution := THValComboBox(GetControl('PE2_TYPEEVOLMSA'));
        {$ENDIF}
        If TypeEvolution <> Nil then TypeEvolution.OnChange := ChangeTypeEvolution;

  //DEB PT4
  Combo := THValComboBox(GetControl('PERIODICITE'));
  Combo.OnChange := cbComboChange;
  Edit := THEdit(GetControl('DUREECTR'));
  Edit.OnChange := cbValChange;
  //FIN PT4
end ;

procedure TOM_MSAEVOLUTIONSPE2.OnClose ;                                     
begin
  Inherited ;
end ;

procedure TOM_MSAEVOLUTIONSPE2.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_MSAEVOLUTIONSPE2.ChangeTypeEvolution (Sender : TObject);
var PContrat,PCompl : TTabSheet;
    Test : String;
begin
        PContrat := TTabSheet(GetControl('PE23'));
        PCompl := TTabSheet(GetControl('PE24'));
        PContrat.TabVisible := False;
        PCompl.TabVisible := False;
        SetControlVisible('GRBXPE21',False);
        SetControlVisible('GRBX22',False);
        Test := GetControlText('PE2_TYPEEVOLMSA');
        If GetControlText('PE2_TYPEEVOLMSA') = 'PE21' then SetControlVisible('GRBXPE21',True);
        If GetControlText('PE2_TYPEEVOLMSA') = 'PE22' then SetControlVisible('GRBX22',True);
        If GetControlText('PE2_TYPEEVOLMSA') = 'PE23' then PContrat.TabVisible := True;
        If GetControlText('PE2_TYPEEVOLMSA') = 'PE24' then PCompl.TabVisible := True;
end;

procedure TOM_MSAEVOLUTIONSPE2.ExitEdit(Sender: TObject);   //PT3
var edit : thedit;
    Salarie:String;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<10) and (isnumeric(edit.text)) then
       begin
       Salarie:=AffectDefautCode(edit,10);
       edit.text:=Salarie;
       SetField('PE2_SALARIE',Salarie);
       end;
end;

//DEB PT4
procedure TOM_MSAEVOLUTIONSPE2.cbComboChange(Sender: TObject);
begin
  SetField('PE2_SALARIE',GetField('PE2_SALARIE'));
end;

procedure TOM_MSAEVOLUTIONSPE2.cbValChange(Sender: TObject);
begin
  SetField('PE2_SALARIE',GetField('PE2_SALARIE'));
end;
//FIN PT4

Initialization
  registerclasses ( [ TOM_MSAEVOLUTIONSPE2 ] ) ;
end.

