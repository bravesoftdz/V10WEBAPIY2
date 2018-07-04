{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 09/02/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTPARAMSOCGEN ()
Mots clefs ... : TOF;BTPARAMSOCGEN
*****************************************************************}
Unit BTPARAMSOCGEN_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
{$ENDIF}
     uTob,
     forms,
     vierge,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UtilLine,
     UTOF ;

const TypeChampsTraitable : string = 'TEdit;THedit;THNumEdit;TCheckBox;THCheckbox;THSpinEdit;THGroupBox;TGroupBox;THValComboBox;THMultiValComboBox;THCritMaskEdit;TPanel;THPanel;TRadioButton';

Type
  TOF_BTPARAMSOCGEN = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  public
    //procedure ChargeEcran;
    //procedure StockeInfos;
  private
    //
    TOBLOC : TOB;
    function RecupListeChamp : string;
    procedure ConstitueLaTOB;
    //procedure ChargeComponent(composant: TWinControl);
    //procedure DescentChargeUnNiveau(Composant: TWinControl);
    //procedure DescentStockeUnNiveau(Composant: TWinControl);
    //procedure StockeComponent(composant: TWinControl);
    function RecupListeChampFille(Composant: TWinControl): string;
  end ;

Implementation

procedure TOF_BTPARAMSOCGEN.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMSOCGEN.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMSOCGEN.OnUpdate ;
begin
  Inherited ;

  if TOBLOC <> nil then
  	 begin
     StockeInfos(Ecran, TOBLOC);
     TOBLOC.InsertOrUpdateDB (true);
     end;

end ;

procedure TOF_BTPARAMSOCGEN.OnLoad ;
begin
  Inherited ;

end ;

procedure TOF_BTPARAMSOCGEN.OnArgument (S : String ) ;
begin
  Inherited ;

  TOBLOC := nil;

  //ALONE = ''
  if S = 'ALONE' then
  	 begin
     ConstitueLaTOB;
     LaTOB := TOBLoc;
  	 end;

end ;

procedure TOF_BTPARAMSOCGEN.OnClose ;
begin
  Inherited ;
  if TOBLOC <> nil then FreeAndNil (TOBLoc);
end ;

procedure TOF_BTPARAMSOCGEN.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMSOCGEN.OnCancel () ;
begin
  Inherited ;
end ;


{procedure TOF_BTPARAMSOCGEN.ChargeComponent (composant : TWinControl);
var NomToFind 	: string;
		TheTOBValue : TOB;
    SpinVal	  	: Integer;
begin

  NomToFind := composant.Name;

  TheTOBValue := LaTob.findFirst(['SOC_NOM'],[NomtoFind],true);

  if TheTobValue = Nil then exit;

  if (composant IS TEdit) or
     (composant IS THEdit) or
		 (composant IS THNumEdit) or
     (composant IS THValComboBox ) or
     (composant IS THMultiValComboBox ) or
     (composant IS THCritMaskEdit ) then
     if (composant is THCritMaskEdit) and
        (ThCritMaskEdit(composant).OpeType = otDate) then
       	SetControltext(composant.name,DateToStr(TheTOBValue.GetValue ('SOC_DATA')))
     else
        SetControltext(composant.name,TheTOBValue.GetValue ('SOC_DATA'))
  else if (composant IS TCheckBox) or
          (composant IS THCheckBox) then
     TCheckBox(composant).Checked := (TheTOBValue.getValue('SOC_DATA')='X')
  else if composant IS THSpinEdit then
     Begin
     SpinVal := StrToInt(TheTOBValue.getValue('SOC_DATA'));
     THSpinEdit(composant).Value := SpinVal;
     end;

end;}

{procedure TOF_BTPARAMSOCGEN.DescentChargeUnNiveau (Composant : TWinControl);
var indice : integer;
		NomToFind : string;
begin

	for indice := 0 to composant.Controlcount -1 do
     begin
  	 NomToFind := Composant.Controls [Indice].Name;
     if (Pos(UpperCase(Composant.Controls [Indice].ClassType.className),UpperCase(TypeChampsTraitable))=0) then continue;
  	 if (TwinControl(composant.controls[indice]).ControlCount > 0 ) and
        (composant.controls[indice].ClassType.className<>'THMultiValComboBox') and
        (composant.controls[indice].ClassType.className<>'THValComboBox') and
        (composant.controls[indice].ClassType.className<>'TCheckBox') and
        (composant.controls[indice].ClassType.className<>'THEdit') and
        (composant.controls[indice].ClassType.className<>'TEdit') and
        (composant.controls[indice].ClassType.className<>'THNumEdit') and
        (composant.controls[indice].ClassType.className<>'THSpinEdit') and
        (composant.controls[indice].ClassType.className<>'THCritMaskEdit') then
    	  DescentChargeUnNiveau (TWinControl(composant.Controls[indice]))
     else
    	  ChargeComponent (TWinControl(composant.Controls[indice]));
     end;

end;}

{procedure TOF_BTPARAMSOCGEN.ChargeEcran ;
var indice : integer;
		NomToFind : string;
begin

	for indice := 0 to ecran.ControlCount -1 do
     begin
  	 NomToFind := ecran.Controls [Indice].Name;
     if (Pos(UpperCase(ecran.controls[indice].ClassType.className),UpperCase(TypeChampsTraitable))=0) then continue;
     if (TWinControl(ecran.controls[indice]).ControlCount  > 0) and
        (ecran.controls[indice].ClassType.className<>'THMultiValComboBox') and
        (ecran.controls[indice].ClassType.className<>'THValComboBox') and
        (ecran.controls[indice].ClassType.className<>'TCheckBox') and
        (ecran.controls[indice].ClassType.className<>'THEdit') and
        (ecran.controls[indice].ClassType.className<>'TEdit') and
        (ecran.controls[indice].ClassType.className<>'THNumEdit') and
        (ecran.controls[indice].ClassType.className<>'THSpinEdit') and
        (ecran.controls[indice].ClassType.className<>'THCritMaskEdit') then
    	  DescentChargeUnNiveau (TWinControl(ecran.controls[indice]))
     else
      	 ChargeComponent (TWinControl(ecran.controls[indice]))
     end;


end;}


{procedure TOF_BTPARAMSOCGEN.StockeComponent (composant : TWinControl);
var	NomToFind 	: String;
		Value				: string;
		TheTOBValue : TOB;
begin

	NomtoFind := composant.name;

  TheTOBValue := LaTOB.findFirst(['SOC_NOM'],[NomtoFind],true);

  if TheTOBValue <> nil then
     begin
     if (Composant IS TEdit) or
        (composant IS THEdit) or
        (composant IS THNumEdit) or
        (Composant IS THMultiValComboBox ) or
        (Composant IS THValComboBox ) or
        (composant IS THCritMaskEdit ) then
        begin
    	  Value := GetControlText(Composant.Name);
        if (composant is THCritMaskEdit) and (ThCritMaskEdit(composant).OpeType = otDate) then
           TheTOBValue.PutValue('SOC_DATA',FloatToStr(StrToDate(Value)))
        else
           TheTOBValue.PutValue('SOC_DATA',Value);
        end
     else if Composant IS TCheckBox then
        if TCheckBox(composant).Checked Then
           TheTOBValue.PutValue('SOC_DATA','X')
        Else
           TheTOBValue.PutValue('SOC_DATA','-')
     else if Composant IS THCheckBox then
        if THCheckBox(Composant).Checked Then
           TheTOBValue.PutValue('SOC_DATA','X')
      	Else
           TheTOBValue.PutValue('SOC_DATA','-')
     else if Composant IS THSpinEdit then
        begin
        Value := IntToStr(THSpinEdit(Composant).value);
        TheTOBValue.PutValue('SOC_DATA',Value)
        end;
     end;

end;

procedure TOF_BTPARAMSOCGEN.DescentStockeUnNiveau (Composant : TWinControl);
var indice : integer;
		NomToFind : string;
begin

	for indice := 0 to composant.Controlcount -1 do
  begin
  	NomToFind := Composant.Controls [Indice].Name;
    if (Pos(UpperCase(Composant.Controls [Indice].ClassType.className),UpperCase(TypeChampsTraitable))=0) then continue;
  	if (TwinControl(composant.controls[indice]).ControlCount > 0) and
       (Composant.controls[indice].ClassType.className<>'THMultiValComboBox') and
       (Composant.controls[indice].ClassType.className<>'THValComboBox') and
       (Composant.controls[indice].ClassType.className<>'TCheckBox') and
       (Composant.controls[indice].ClassType.className<>'THEdit') and
       (Composant.controls[indice].ClassType.className<>'TEdit') and
			 (Composant.controls[indice].ClassType.className<>'THNumEdit') and
       (Composant.controls[indice].ClassType.className<>'THSpinEdit') and
       (Composant.controls[indice].ClassType.className<>'THCritMaskEdit') then
    begin
    	DescentStockeUnNiveau (TWinControl(composant.Controls[indice]));
    end else
    begin
    	StockeComponent (TWinControl(composant.Controls[indice]));
    end;
  end;

end;

procedure StockeInfos;
var indice 		: integer;
		NomToFind : string;
begin

	for indice := 0 to ecran.ControlCount -1 do
  		begin
    	NomToFind := ecran.Controls [Indice].Name;
      if (Pos(upperCase(ecran.controls[indice].ClassType.className),UpperCase(TypeChampsTraitable))=0) then continue;
      if (TWinControl(ecran.controls[indice]).ControlCount  > 0) and
         (ecran.controls[indice].ClassType.className<>'THMultiValComboBox') and
         (ecran.controls[indice].ClassType.className<>'THValComboBox') and
         (ecran.controls[indice].ClassType.className<>'TCheckBox') and
         (ecran.controls[indice].ClassType.className<>'THEdit') and
         (ecran.controls[indice].ClassType.className<>'THNumEdit') and
         (ecran.controls[indice].ClassType.className<>'TEdit') and
         (ecran.controls[indice].ClassType.className<>'THSpinEdit') and
         (ecran.controls[indice].ClassType.className<>'THCritMaskEdit') then
		     DescentStockeUnNiveau (TWinControl(ecran.controls[indice]));
    else
         StockeComponent(TWinControl(ecran.controls[indice]));
  end;

end;}

procedure TOF_BTPARAMSOCGEN.ConstitueLaTOB;
var LesChamps : string;
    ZoneTempo : String;
    Tempo			: String;
begin

  TOBLOC := TOB.Create ('TOB LOCALE PARAMSOC',nil,-1);

  LesChamps := RecupListeChamp;

  //Récupération valeur de argument
  ZoneTempo := (Trim(ReadTokenSt(LesChamps)));

  while (ZoneTempo <> '') do
      begin
      if Tempo = '' then
         Tempo := '"' + ZoneTempo + '"'
      else
	       Tempo := Tempo + ',"' + ZoneTempo + '"';
      ZoneTempo := (Trim(ReadTokenSt(LesChamps)));
    end;

  LesChamps := Tempo;

  if lesChamps = '' then
	   TOBLoc.LoadDetailDBFromSQL ('PARAMSOC','SELECT * FROM PARAMSOC',false,true)
  else
	   TOBLoc.LoadDetailDBFromSQL ('PARAMSOC','SELECT * FROM PARAMSOC WHERE SOC_NOM IN ('+lesChamps+')',false,true);

end;


function TOF_BTPARAMSOCGEN.RecupListeChampFille (Composant : TWinControl) : string;
var indice : integer;
		NomToFind : string;
    valloc : string;
begin
  result := '';

	for indice := 0 to composant.Controlcount -1 do
  begin
  	NomToFind := Composant.Controls [Indice].Name;
    if (Pos(UpperCase(Composant.Controls [Indice].ClassType.className),UpperCase(TypeChampsTraitable))=0) then continue;
  	if (TwinControl(composant.controls[indice]).ControlCount > 0 ) and
       (composant.controls[indice].ClassType.className<>'THMultiValComboBox') and
       (composant.controls[indice].ClassType.className<>'THValComboBox') and
       (composant.controls[indice].ClassType.className<>'TCheckBox') and
       (composant.controls[indice].ClassType.className<>'THEdit') and
			 (composant.controls[indice].ClassType.className<>'THNumEdit') and
       (composant.controls[indice].ClassType.className<>'TEdit') and
       (composant.controls[indice].ClassType.className<>'THSpinEdit') and
       (composant.controls[indice].ClassType.className<>'THCritMaskEdit') then
    begin
    	 Valloc := RecupListeChampFille (TWinControl(ecran.controls[indice]));
       if result = '' then
       		result := Valloc
       else
       		result := result + ';' + Valloc;
    end
    else
    begin
       if result = '' then
      	  result := NomToFind
       else
          result := result + ';' + NomToFind;
    end;
  end;

end;

function TOF_BTPARAMSOCGEN.RecupListeChamp: string;
var indice : integer;
    valloc : string;
		NomToFind : string;
begin
  result := '';

	for indice := 0 to ecran.ControlCount -1 do
  begin
  	NomToFind := ecran.Controls [Indice].Name;
    if (Pos(Uppercase(ecran.controls[indice].ClassType.className),UpperCase(TypeChampsTraitable))=0) then continue;
    if (TWinControl(ecran.controls[indice]).ControlCount  > 0) and
       (ecran.controls[indice].ClassType.className<>'THMultiValComboBox') and
       (ecran.controls[indice].ClassType.className<>'THValComboBox') and
       (ecran.controls[indice].ClassType.className<>'TCheckBox') and
       (ecran.controls[indice].ClassType.className<>'THEdit') and
       (ecran.controls[indice].ClassType.className<>'TEdit') and
			 (ecran.controls[indice].ClassType.className<>'THNumEdit') and
       (ecran.controls[indice].ClassType.className<>'THSpinEdit') and
       (ecran.controls[indice].ClassType.className<>'THCritMaskEdit') then
       begin
    	 Valloc := RecupListeChampFille (TWinControl(ecran.controls[indice]));
       if result = '' then
          result := Valloc
       else
          result := result + ';'+ Valloc;
       end
    else
    begin
       if result = '' then
          result := NomToFind
       else
          result := result + ';' + NomToFind;
    end;
  end;


end;

Initialization
  registerclasses ( [ TOF_BTPARAMSOCGEN ] ) ;
end.

