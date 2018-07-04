{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 04/07/2002
Modifié le ... : 04/07/2002
Description .. : Source TOM de la TABLE : RELANCE (RELANCE)
Mots clefs ... : TOM;RELANCE
*****************************************************************}
Unit RELANCE_TOM ;

Interface

Uses Windows, StdCtrls, Controls, Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
     Fe_main, Fiche, FichList,
{$ELSE}
		MaineAGL, eFiche, eFichList, 
{$ENDIF}
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOM, UTob, Spin;

procedure CCLanceFiche_ParamRelance(pszType : String);

Type
  TOM_RELANCE = Class (TOM)
    procedure OnNewRecord                ; override ;
	  procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ; // BVE 31.05.07 FQ 20470
    procedure OnLoadRecord               ; override ;
    procedure OnArgument ( S: String )   ; override ;
  private
		gszTitre : String;
  	procedure RR_SCOORINGClick(Sender: TObject);
    procedure RR_GROUPELETTREClick(Sender: TObject);
    procedure RR_MODELEKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    function  VerifJour : Boolean;
    function  ContinuiteDelaiModele : Boolean ;
end ;

Const	MsgBox: array[0..8] of string = (
	'Paramètres de relance des règlements',
	'Paramètres de relance des traites',
	'2;?caption?;$$ possède un nombre de jours inférieur ou égal au niveau précédent !;E;O;O;O;',
	'3;',
	'Le niveau de relance N°',
	'5;?caption?;Dès lors qu''un modèle est renseigné, vous devez spécifier un délai.;W;O;O;O;',
	'6;?caption?;Si un délai est renseigné vous devez spécifier un modèle.;E;O;O;O;',
	'7;?caption?;Vous devez renseigner au moins un délai !;E;O;O;O;',
	'8;?caption?;Il n''y a pas de continuité dans la saisie délai,modèle !;E;O;O;O;');

Implementation
uses Ent1; //ChargeTableLibre

procedure CCLanceFiche_ParamRelance(pszType : String);
begin
	AGLLanceFiche('CP','CPPARAMREL',pszType,'',pszType);
end;

procedure TOM_RELANCE.OnUpdateRecord ;
begin
  Inherited ;

  //Si Erreur
  if Not ContinuiteDelaiModele then begin LastError := 1; Exit; end;
  if Not VerifJour then begin LastError := 1; Exit; end;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 31/05/2007
Modifié le ... : 31/05/2007
Description .. : Permet de mettre à jour la tablette des relances
Suite ........ : FQ 20470
Mots clefs ... : 
*****************************************************************}
procedure TOM_RELANCE.OnAfterUpdateRecord ;
begin
  Inherited ;
{$IFDEF EAGLCLIENT}
  AvertirCacheServer('RELANCE');
{$ELSE}
  ChargeTableLibre;
{$ENDIF}
  AvertirTable('TTRELANCEREGLE');
end ;


procedure TOM_RELANCE.OnArgument ( S: String ) ;
var
  cControl : TCheckBox;
  cControlb: TComboBox;
  i : integer;
//{$IFDEF CCS3}
//	i : integer;
//{$ENDIF}
begin
  Inherited ;
  cControl := TCheckBox(GetControl('RR_SCOORING')); if (cControl <> nil) then cControl.OnClick := RR_SCOORINGClick;
  cControl := TCheckBox(GetControl('RR_GROUPELETTRE')); if (cControl <> nil) then cControl.OnClick := RR_GROUPELETTREClick;
	for i:=1 to 7 do begin
  	cControlb := TComboBox(GetControl('RR_MODELE'+IntToStr(i))); if (cControlb <> nil) then cControlb.OnKeyDown := RR_MODELEKeyDown;
  end;

	if S ='RTR' then begin
  	Ecran.HelpContext := 7574000;
    gszTitre := MsgBox[1];
    end
  else begin
  	Ecran.HelpContext := 7577000;
    gszTitre := MsgBox[0];
  end;
  UpdateCaption(Ecran);

	{$IFDEF CCS3}
 { YMO 09/11/2005 FQ 16778
	for i:=4 to 7 do begin
    TWinControl(GetControl('TRR_DELAI'+IntToStr(i))).Visible := False;
    TWinControl(GetControl('RR_DELAI'+IntToStr(i))).Visible := False;
    TWinControl(GetControl('LBL_JOUR'+IntToStr(i))).Visible := False;
    TWinControl(GetControl('RR_MODELE'+IntToStr(i))).Visible := False;
	end;
  }
	TWinControl(GetControl('RR_SCOORING')).Visible := False;
	{$ENDIF}

    SetControlVisible('BImprimer',false);
end ;

procedure TOM_RELANCE.RR_GROUPELETTREClick(Sender: TObject);
begin
if TCheckBox(GetControl('RR_GROUPELETTRE')).Checked then SetControlChecked('RR_SCOORING',False);
end;

procedure TOM_RELANCE.RR_SCOORINGClick(Sender: TObject);
begin
if TCheckBox(GetControl('RR_SCOORING')).Checked then SetControlChecked('RR_GROUPELETTRE',False);
end;

function TOM_RELANCE.VerifJour: Boolean;
var
	Model : String ;
  Delai,i : Integer ;
begin
	Result:=False ;
	for i:=1 to 7 do begin
	  Model:=GetControlText('RR_MODELE'+IntToStr(i));
  	Delai:=TSpinEdit(GetControl('RR_DELAI'+IntToStr(i))).Value;
	  if(Delai=0) And (Model<>'') then begin
	    If i<>1 then begin
      	HShowMessage(MsgBox[5],'','') ; //Dès lors qu'un modèle est renseigné, vous devez spécifier un délai
      	SetFocusControl('RR_DELAI'+IntToStr(i)); Exit;
      end;
    	end
    else if(Delai<>0) And (Model='') then begin
    	HShowMessage(MsgBox[6],'','') ; //Si un délai est renseigné vous devez spécifier un modèle
      SetFocusControl('RR_MODELE'+IntToStr(i)); Exit;
      end
    else if i>1 then begin
    	if(Delai<>0) And (Model<>'') then begin
      	if Delai<=TSpinEdit(GetControl('RR_DELAI'+IntToStr(i-1))).Value then begin
        	HShowMessage(MsgBox[2],MsgBox[4]+' '+IntToStr(i)+' ','') ; // Le niveau de relance N° ? possède un nombre de jours inférieur ou égal au niveau précédent !
          SetFocusControl('RR_DELAI'+IntToStr(i)); Exit;
        end;
      end;
    end;
	end;
	Result:=True ;
end;

function TOM_RELANCE.ContinuiteDelaiModele: Boolean;
var
	i : Byte ;
  Trou : Boolean ;
begin
	Result:=False ; Trou:=True ;
	for i:=1 to 7 do
  	if TSpinEdit(GetControl('RR_DELAI'+IntToStr(i))).Value<>0 then begin
    	Trou:=False ;
      Break ;
    end;
	if Trou then begin
		HShowMessage(MsgBox[7],'','') ; //Vous devez renseigner au moins un délai !
		SetFocusControl('RR_DELAI1'); exit;
	end;

  for i:=7 Downto 2 do begin
  	if(TSpinEdit(GetControl('RR_DELAI'+IntToStr(i))).Value<>0) And (GetControlText('RR_MODELE'+IntToStr(i))<>'') then
    	if(TSpinEdit(GetControl('RR_DELAI'+IntToStr(i-1))).Value=0) And (GetControlText('RR_MODELE'+IntToStr(i-1))='') then begin
      	HShowMessage(MsgBox[8],'','') ; //Il n'y a pas de continuité dans la saisie délai,modèle !
       	SetFocusControl('RR_DELAI'+IntToStr(i-1));
        Exit ;
      end;
  end;
	Result:=True ;
end;

procedure TOM_RELANCE.OnNewRecord;
begin
  inherited;
	Ecran.caption := gszTitre;
  UpdateCaption(Ecran);
end;

procedure TOM_RELANCE.OnLoadRecord;
begin
  inherited;
	Ecran.caption := gszTitre+' '+GetControlText('RR_FAMILLERELANCE')+' '+GetControlText('RR_LIBELLE');
	UpdateCaption(Ecran);
end;

procedure TOM_RELANCE.RR_MODELEKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
	IF Key = VK_DELETE then SetControlText(TControl(Sender).Name,'');
end;

Initialization
  registerclasses ( [ TOM_RELANCE ] ) ;
end.
