{***********UNITE*************************************************
Auteur  ...... : CATALA David
Créé le ...... : 27/03/2007
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : DPPARCELLES (DPPARCELLES)
Mots clefs ... : TOM;DPPARCELLES
*****************************************************************}
Unit DPPARCELLES_TOM ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fiche,
     FichList, 
{$else}
     eFiche,
     eFichList,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOM,
     UTob ;

Type
  TOM_DPPARCELLES = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnArgument ( S: String )   ; override ;

    procedure ControlOnChangeDroitsUtilisation (Sender: TObject);
    procedure ControlOnClickEnCulture (Sender: TObject);
    function  IsCbNivCocher (CbNiv : String) : Boolean;
    function ConstruireListeCodeCulture (UnNiveau : String) : String;

   private
    NumDossier     : String;
    CulturePerenne : String;
    ListeCode      : String;
  end ;

Implementation

Const
 ListeCbNiv21 : Array [1..6]  of String =('DAG_101CEREALE','DAG_102OLEOPROTE','DAG_103FOURRAGE','DAG_104LEGUME','DAG_105FLORALE','DAG_106AUTRE');
 ListecbNiv22 : Array [1..5]  of String =('DAG_201FRUITIERE','DAG_202VITICULTURE','DAG_203SYLVICULT','DAG_204PEPINIERE','DAG_205AUTRE');

//----------------------------------------
//--- Nom : OnArgument
//----------------------------------------
procedure TOM_DPPARCELLES.OnArgument ( S: String ) ;
var sFiltreListe : String;
    sListeCbNiv  : String;
    UneListeCode : String;
    Indice       : Integer;
begin
 Inherited ;
 //--- Activation des controls
 THValComboBox (GetControl ('DPC_DROITPARCELLE')).OnChange:=ControlOnChangeDroitsUtilisation;
 THCheckBox (GetControl ('DPC_ENCULTURE')).Onclick:=ControlOnClickEnCulture;

 //--- Récupère les paramètres de la fiche
 NumDossier:=ReadToKenSt (S);
 CulturePerenne:=ReadToKenSt (S);
 ListeCode:=ReadToKenSt (S);

 //--- Construction du filtre
 sFiltreListe:= sFiltreListe+'DPC_NODOSSIER='''+NumDossier+''' AND DPC_CULTPERENNE='''+CulturePerenne+'''';

 {$IFDEF EAGLCLIENT}
  sFiltreListe:=Trim(StringReplace(sFiltreListe,'''','"',[rfReplaceAll, rfIgnoreCase]));
  TFFicheListe(Ecran).SetNewRange('',sFiltreListe);
 {$ELSE}
  TFFicheListe(Ecran).Ta.Filter:=sFiltreListe;
  TFFicheListe(Ecran).Ta.Filtered:=True;
 {$ENDIF}

 //--- Mise à jour de la propriété PLUS De la combo DPC_CULTURE
 SListeCbNiv:='';
 if (StrToBool (CulturePerenne)) then
  begin
   for Indice:=1 to 5 do
    if (IsCbNivCocher (ListeCbNiv22 [Indice])) then
     begin
      UneListeCode:=ConstruireListeCodeCulture (Copy (ListeCbNiv22 [Indice],5,3));
      if (UneListeCode<>'') then
       begin
        if (SListeCbNiv='') then
         SListeCbNiv:=SListeCbNiv+UneListeCode
        else
         SListeCbNiv:=SListeCbNiv+' OR '+UneListeCode;
       end;
     end;
  end
 else
  begin
   for Indice:=1 to 6 do
    if (IsCbNivCocher (ListeCbNiv21 [Indice])) then
     begin
      UneListeCode:=ConstruireListeCodeCulture (Copy (ListeCbNiv21 [Indice],5,3));
      if (UneListeCode<>'') then
       begin
        if (SListeCbNiv='') then
         SListeCbNiv:=SListeCbNiv+UneListeCode
        else
         SListeCbNiv:=SListeCbNiv+' OR '+UneListeCode;
       end;
     end;
  end;

 if (SListeCbNiv<>'') then
  THValCombobox (GetControl ('DPC_CULTURE')).plus:=' AND ('+SListeCbNiv+')';
end ;

//----------------------------------------
//--- Nom : ConstruireListeCode
//----------------------------------------
function TOM_DPPARCELLES.ConstruireListeCodeCulture (UnNiveau : String) : String;
var UnCode : String;
    UneListeCode : String;
    SSql : String;
begin
 SSql:='';
 UneListeCode:=ListeCode;
 While (UneListeCode<>'') do
  begin
  UnCode:=ReadToKenPipe (UneListeCode,'|');
   if (Copy (UnCode,1,3)=UnNiveau) then
    begin
     if (SSql='') then
      SSql:=' YDS_CODE="'+UnCode+'"'
     else
      SSql:=SSql+' OR YDS_CODE="'+UnCode+'"'
    end;
  end;
 Result:=SSql;
end;

//----------------------------------------
//--- Nom : OnNewRecord
//----------------------------------------
procedure TOM_DPPARCELLES.OnNewRecord;
begin
 Inherited ;
 SetField ('DPC_NODOSSIER',NumDossier);
 SetField ('DPC_GUIDPARCELLE',AglGetGuid());
 SetField ('DPC_CULTPERENNE',CulturePerenne);
 THCheckBox (GetControl ('DPC_ENCULTURE')).Checked:=False;
end ;

//----------------------------------------
//--- Nom : OnLoadRecord
//----------------------------------------
procedure TOM_DPPARCELLES.OnLoadRecord;
begin
 Inherited;
 if (StrToBool (CulturePerenne)) then
  THLabel (GetControl('LTYPECULTURE')).caption:='Gestion cultures pérennes'
 else
  THLabel (GetControl('LTYPECULTURE')).caption:='Gestion cultures avec cycle annuel';
 UpdateCaption(Ecran);

 ControlOnChangeDroitsUtilisation(Nil);
 ControlOnClickEnCulture(Nil);
end;

//-----------------------------------------
//--- Nom : ControlOnExitCultureAnnuelle
//-----------------------------------------
procedure TOM_DPPARCELLES.ControlOnChangeDroitsUtilisation (Sender: TObject);
begin
 if (THValComboBox (GetControl('DPC_DROITPARCELLE')).Text='Exploitation directe') or (THValComboBox (GetControl('DPC_DROITPARCELLE')).Text='Prise à bail') then
  begin
   THCheckBox (GetControl('DPC_ENCULTURE')).Enabled:=True;
   THValComboBox (GetControl('DPC_CULTURE')).Enabled:=THCheckBox (GetControl('DPC_ENCULTURE')).Checked;
   THLabel (GetControl('TDPC_CULTURE')).Enabled:=THCheckBox (GetControl('DPC_ENCULTURE')).Checked;
  end
 else
  begin
   THCheckBox (GetControl('DPC_ENCULTURE')).Enabled:=False;
   if THCheckBox (GetControl ('DPC_ENCULTURE')).Checked<>False then THCheckBox (GetControl ('DPC_ENCULTURE')).Checked:=False;
  end;

 if (THValComboBox (GetControl('DPC_DROITPARCELLE')).Text='Prise à bail') then
  begin
   THValCombobox (GetControl ('DPC_MODEPARCELLE')).plus:='AND CO_CODE<>"PRO"';
   THValComboBox (GetControl('DPC_MODEPARCELLE')).Enabled:=True;
   THLabel (GetControl('TDPC_MODEPARCELLE')).Enabled:=True;
  end
 else
  begin
   THValCombobox (GetControl ('DPC_MODEPARCELLE')).plus:='AND CO_CODE="PRO"';
   if (THValComboBox (GetControl('DPC_MODEPARCELLE')).Value<>'PRO') then
    THValComboBox (GetControl('DPC_MODEPARCELLE')).Value:='PRO';
   THValComboBox (GetControl('DPC_MODEPARCELLE')).Enabled:=False;
  end;
end;

//------------------------------------
//--- Nom : ControlOnClickEnCulture
//------------------------------------
procedure TOM_DPPARCELLES.ControlOnClickEnCulture (Sender: TObject);
var Etat : Boolean;
begin
 Etat:=THCheckBox (GetControl ('DPC_ENCULTURE')).Checked;
 THValComboBox (GetControl('DPC_CULTURE')).Enabled:=Etat;
 THLabel (GetControl('TDPC_CULTURE')).Enabled:=Etat;
end;

//--------------------------
//--- Nom : IsCbNivCocher
//--------------------------
function TOM_DPPARCELLES.IsCbNivCocher (CbNiv : String) : Boolean;
var SSql     : String;
    RSql     : TQuery;
begin
 Result:=False;
 SSql:='SELECT '+CbNiv+' FROM DPAGRICOLE WHERE DAG_NODOSSIER="'+NumDossier+'"';
 RSql:=OpenSql (SSql, True);
 if (not RSql.Eof) then
  Result:=StrToBool (RSql.FindField (CbNiv).AsString);
 Ferme (RSql);
end;

Initialization
 registerclasses ( [ TOM_DPPARCELLES ] ) ;
end.

