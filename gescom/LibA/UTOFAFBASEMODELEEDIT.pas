{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 20/10/2000
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : AFBASEMODELEEDIT ()
Mots clefs ... : TOF;AFPROPOSEDITMUL
*****************************************************************}
Unit UTOFAFBASEMODELEEDIT ;

Interface

Uses
{$IFDEF EAGLCLIENT}
     eMul,
{$ELSE}
     Mul,EdtDoc ,db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,
{$ENDIF}
     StdCtrls, Controls, Classes, forms, sysutils, ComCtrls, HTB97,
     HCtrls, HEnt1, HMsgBox, UTOF, Dicobtp, UTofAfBaseCodeAffaire, paramsoc,
     TraducAffaire ;

Type
  TOF_AFBASEMODELEEDIT = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnArgument (stArgument : String ) ; override ;
    procedure RBEditDocClick(Sender: TObject);
    procedure RBEditEtatClick(Sender: TObject);
    public
    Statut : string;
    DocPro ,EtatPro ,DocAff ,EtatAff : string;
    ModeleEdit : THValComboBox;
    END ;

//const
// libellés des messages
//TexteMessage: array[1..3] of string 	= (
//          {1}  'Veuillez sélectionner un modèle d''édition'
//          {2} ,'Veuillez sélectionner au-moins une '
//          {3} ,' dans la liste'
//              );

implementation


procedure TOF_AFBASEMODELEEDIT.OnArgument (stArgument : String ) ;
var  Critere,champ, valeur, value : string;
x:integer;
begin
  Inherited ;
ModeleEdit := THValComboBox(GetControl('MODELEEDITION'));
if (ModeleEdit=nil) then exit;
DocPro := '';
EtatPro := '';
DocAff := '';
EtatAff := '' ;
value:='';


Statut := 'AFF';
Critere:=(Trim(ReadTokenSt(stArgument)));
While (Critere <>'') do
    BEGIN
    X:=pos('=',Critere);
    if x<>0 then
           begin
           Champ:=copy(Critere,1,X-1);
           Valeur:=Copy (Critere,X+1,length(Critere)-X);
           end;

    if (Champ='STATUT') then
	        begin
            TRadioButton(GetControl('RBEDITDOC')).Checked := true;
            ModeleEdit.DataType := 'AFMODELEAFFAIRE';
            ModeleEdit.value := '';

            if (Valeur='PRO') then
                begin
                Statut := 'PRO';
                DocPro := GetParamSoc('SO_AFDOCUMENTPRO');
                EtatPro := GetParamSoc('SO_AFETATPRO');
                if (EtatPro<>'') then
                    begin
                    ModeleEdit.DataType := 'AFMODELEETATAFFAIRE';
                    ModeleEdit.plus := 'APE';
                    Value := EtatPro;
                    TRadioButton(GetControl('RBEDITETAT')).Checked := true;
                    end
                else
                if (DocPro<>'') then
                    begin
                    ModeleEdit.plus := 'ADE';
                    Value := DocPro;
                    end
                else
                    ModeleEdit.plus := 'ADE';
                end;

            if (Valeur='AFF') then
                begin
                Statut := 'AFF';
                DocAff := GetParamSoc('SO_AFDOCUMENTAFF');
                EtatAff := GetParamSoc('SO_AFETATAFF');
                if (EtatAff<>'') then
                    begin
                    ModeleEdit.DataType := 'AFMODELEETATAFFAIRE';
                    ModeleEdit.plus := 'AFE';
                    Value := EtatAff;
                    TRadioButton(GetControl('RBEDITETAT')).Checked := true;
                    end
                else
                if (DocAff<>'') then
                    begin
                    ModeleEdit.plus := 'AFF';
                    Value := DocAff;
                    end
                else
                    ModeleEdit.plus := 'AFF';
                end;

            if (Valeur='FAC') then
                begin
                                    // mcd 14/12/01 etat en priorité, document plus maintenu
                TRadioButton(GetControl('RBEDITDOC')).Checked := false;
                TRadioButton(GetControl('RBEDITETAT')).Checked := true;
                ModeleEdit.DataType := 'GCIMPETAT';
                Statut := 'FAC';
                end;
	        end;

    Critere:=(Trim(ReadTokenSt(stArgument)));
    END;

  ModeleEdit.Refresh;
  ModeleEdit.items.Commatext :=TraduitGA(ModeleEdit.items.Commatext);
  ModeleEdit.value := Value;
  TRadioButton(GetControl('RBEDITDOC')).OnClick := RBEditDocClick;
  TRadioButton(GetControl('RBEDITETAT')).OnClick := RBEditEtatClick;

end ;

procedure TOF_AFBASEMODELEEDIT.RBEditDocClick(Sender: TObject);
var
bModif:boolean;
Value:string;
begin
Value:='';
bModif:=false;

if (TRadioButton(GetControl('RBEDITDOC')).Checked = true) then
begin
    ModeleEdit.DataType := 'AFMODELEAFFAIRE';
    if (Statut='PRO') then
        begin
             bModif := (ModeleEdit.plus <> 'ADE');
             if bModif then
             begin
              ModeleEdit.plus := 'ADE';
              Value := DocPro;
             end;
        end;

    if (Statut='AFF') then
    begin
        bModif := (ModeleEdit.plus <> 'AFF');
        if bModif then
        begin
            ModeleEdit.plus := 'AFF';
            Value := DocAff;
        end;
    end;

    if (Statut='FAC') then
        begin
        ModeleEdit.DataType := 'GCIMPMODELE';
        bModif := (ModeleEdit.plus <> 'GPI');
        if bModif then
            begin
            Value :='';
//            ModeleEdit.plus := 'GPI';
            end;

        end;
    end
else
    begin
    end;

if bModif then
    begin
    ModeleEdit.Refresh;
    ModeleEdit.items.Commatext :=TraduitGA(ModeleEdit.items.Commatext);
    ModeleEdit.Value := Value;
    end;
end;

procedure TOF_AFBASEMODELEEDIT.RBEditEtatClick(Sender: TObject);
var
bModif:boolean;
Value:string;
begin
Value:='';
bModif:=false;

if (TRadioButton(GetControl('RBEDITETAT')).Checked = true) then
    begin
    ModeleEdit.DataType := 'AFMODELEETATAFFAIRE';

    if (Statut='PRO') then
        begin
        bModif := (ModeleEdit.plus <> 'APE');
        if bModif then
            begin
            ModeleEdit.plus := 'APE';
            Value := EtatPro;
            end;
        end;

    if (Statut='AFF') then
        begin
        bModif := (ModeleEdit.plus <> 'AFE');
        if bModif then
            begin
            ModeleEdit.plus := 'AFE';
            Value := EtatAff;
            end;
        end;

    if (Statut='FAC') then
        begin
        ModeleEdit.DataType := 'GCIMPETAT';    
        bModif := (ModeleEdit.plus <> 'GPJ');
        if bModif then
            begin
            Value :='';
//            ModeleEdit.plus := 'GPJ';
            end;
        end;
    end
else
    begin
    end;

if bModif then
    begin
    ModeleEdit.Refresh;
    ModeleEdit.items.Commatext :=TraduitGA(ModeleEdit.items.Commatext);
    ModeleEdit.Value := value;
    end;
end;

Initialization
registerclasses ([TOF_AFBASEMODELEEDIT]) ;
end.
