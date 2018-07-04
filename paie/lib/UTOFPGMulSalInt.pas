{***********UNITE*************************************************
Auteur  ...... : FLO
Créé le ...... : 26/03/2008
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : INTERIMINSC_MUL ()
Mots clefs ... : TOF;INTERIMINSC_MUL
*****************************************************************
PT1 | 23/04/2008 | V_804 | FL | Gestion de la fiche sans partage (table salarié utilisée à la place de Intérimaires)
PT2 | 04/06/2008 | V_804 | FL | FQ 15458 Ne pas voir les salariés confidentiels (la table PSI n'étant pas gérée en automatique)
}
Unit UtofPGMulSalInt;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     DBGrids, HDB, db, Mul, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
{$ELSE}
      MaineAgl,eMul,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,
     UTOF,
     PGOutilsFormation,
     hTB97;

Type
	TOF_PGMULSALINT = Class (TOF)
		procedure OnUpdate                 ; override ;
		procedure OnLoad                   ; override ;
		procedure OnArgument (S : String ) ; override ;
		private
		{$IFNDEF EAGLCLIENT}
		Liste : THDBGrid;
		{$ELSE}
		Liste : THGrid;
		{$ENDIF}
		Req : String;
        TypeSaisie : String;
        Pfx        : String; //PT1
		procedure GrilleDBLClick(Sender : TObject);
		Procedure ValideSelection (Sender : TObject);
		procedure ClickSalarie (Sender : TObject);
	end ;

Implementation

Uses EntPaie,P5Def,ParamSoc;

procedure TOF_PGMULSALINT.OnUpdate ;
begin
  Inherited ;

end ;

procedure TOF_PGMULSALINT.OnLoad ;
var DateArret : TdateTime;
    Chaine    : String;
begin
	Chaine      := '';
	
	if  TCheckBox(GetControl('CKSORTIE'))<>nil then
	Begin
		if (GetControlText('CKSORTIE')='X') and (IsValidDate(GetControlText('DATEARRET')))then
		Begin
			DateArret:=StrtoDate(GetControlText('DATEARRET'));
			Chaine := ' AND ('+Pfx+'_DATESORTIE>="'+UsDateTime(DateArret)+'" OR '+Pfx+'_DATESORTIE="'+UsdateTime(Idate1900)+'" OR '+Pfx+'_DATESORTIE IS NULL) '; //PT1
			Chaine := Chaine + ' AND '+Pfx+'_DATEENTREE <="'+UsDateTime(DateArret)+'"';
		End;
	End;

	If (GetControlText('CKDEJAINSCRITS')='X') And (TypeSaisie = 'SESSION') Then
	Begin
		If Pfx = 'PSI' Then //PT1
			Chaine := Chaine + ' AND PSI_INTERIMAIRE '
		Else
			Chaine := Chaine + ' AND PSA_SALARIE ';
		Chaine := Chaine + 'NOT IN (SELECT PFO_SALARIE FROM FORMATIONS WHERE '+Req+' AND PFO_ETATINSCFOR<>"REP")';
	End
	Else If (GetControlText('CKDEJAINSCRITS')='X') And (TypeSaisie = 'PREVISIONNEL') Then //PT1
	Begin
		If Pfx = 'PSI' Then //PT1
			Chaine := Chaine + ' AND PSI_INTERIMAIRE '
		Else
			Chaine := Chaine + ' AND PSA_SALARIE ';
		Chaine := Chaine + 'NOT IN (SELECT PFI_SALARIE FROM INSCFORMATION WHERE '+Req+')';
	End;
	
	//PT2
	// Ne pas voir les salariés confidentiels
    If Pfx = 'PSI' Then
    	Chaine := Chaine + ' AND PSI_CONFIDENTIEL<="'+RendNivAccesConfidentiel()+'"'
    Else
    	Chaine := Chaine + ' AND PSA_CONFIDENTIEL<="'+RendNivAccesConfidentiel()+'"';
	
	SetControlText('XX_WHERE', Chaine);
end ;

procedure TOF_PGMULSALINT.OnArgument (S : String ) ;
var
  Num    : Integer;    
  BValid : TToolBarButton97;
  Edit   : THEdit;
  Combo  : THValComboBox;
  Labl  : THLabel;
begin
	TypeSaisie := ReadTokenPipe(S,';');
	
	//PT1
	If PGBundleHierarchie Or GetParamSocSecur('SO_PGINTERVENANTEXT', FALSE) Then
		Pfx := 'PSI'
	Else
	Begin
		Pfx := 'PSA';
		// Changement de la liste
		TFMul(Ecran).SetDBListe('PGSALARIES');

		// Changement des critères
		Edit := THEdit(GetControl('PSI_INTERIMAIRE'));   If Edit <> Nil Then Edit.Name := 'PSA_SALARIE';
		Labl := THLabel(GetControl('TPSI_INTERIMAIRE')); If Labl <> Nil Then Labl.Name := 'TPSA_SALARIE'; Edit.DataType := 'PGSALARIE';
		
		Edit := THEdit(GetControl('PSI_LIBELLE'));   	 If Edit <> Nil Then Edit.Name := 'PSA_LIBELLE';
		Labl := THLabel(GetControl('TPSI_LIBELLE'));   	 If Labl <> Nil Then Labl.Name := 'TPSA_LIBELLE';
		
		Combo:= THValComboBox(GetControl('PSI_LIBELLEEMPLOI')); If Combo <> Nil Then Combo.Name := 'PSA_LIBELLEEMPLOI';
		Labl := THLabel(GetControl('TPSI_LIBELLEEMPLOI')); If Labl <> Nil Then Labl.Name := 'TPSA_LIBELLEEMPLOI';
		
		Combo:= THValComboBox(GetControl('PSI_TRAVAILN1')); If Combo <> Nil Then Combo.Name := 'PSA_TRAVAILN1';
		Labl := THLabel(GetControl('TPSI_TRAVAILN1')); 	 If Labl <> Nil Then Labl.Name := 'TPSA_TRAVAILN1';
		
		Combo:= THValComboBox(GetControl('PSI_TRAVAILN2')); If Combo <> Nil Then Combo.Name := 'PSA_TRAVAILN2';
		Labl := THLabel(GetControl('TPSI_TRAVAILN2')); 	 If Labl <> Nil Then Labl.Name := 'TPSA_TRAVAILN2';
		
		Combo:= THValComboBox(GetControl('PSI_TRAVAILN3')); If Combo <> Nil Then Combo.Name := 'PSA_TRAVAILN3';
		Labl := THLabel(GetControl('TPSI_TRAVAILN3')); 	 If Labl <> Nil Then Labl.Name := 'TPSA_TRAVAILN3';
		
		Combo:= THValComboBox(GetControl('PSI_TRAVAILN4')); If Combo <> Nil Then Combo.Name := 'PSA_TRAVAILN4';
		Labl := THLabel(GetControl('TPSI_TRAVAILN4')); 	 If Labl <> Nil Then Labl.Name := 'TPSA_TRAVAILN4';
		
		Combo:= THValComboBox(GetControl('PSI_CODESTAT'));  If Combo <> Nil Then Combo.Name := 'PSA_CODESTAT';
		Labl := THLabel(GetControl('TPSI_CODESTAT'));  	 If Labl <> Nil Then Labl.Name := 'TPSA_CODESTAT';
	End;

	SetControlText('DATEARRET',DateToStr(V_PGI.DateEntree));
	For Num := 1 to VH_Paie.PGNbreStatOrg do
	begin
		if Num >4 then Break;
		VisibiliteChampSalarie (IntToStr(Num),GetControl (Pfx+'_TRAVAILN'+IntToStr(Num)),GetControl ('T'+Pfx+'_TRAVAILN'+IntToStr(Num)));
	end;
	VisibiliteStat (GetControl (Pfx+'_CODESTAT'),GetControl ('T'+Pfx+'_CODESTAT')) ;
	
	{$IFNDEF EAGLCLIENT}
	Liste := THDBGrid(GetControl('FLISTE'));
	{$ELSE}
	Liste := THGrid(GetControl('FLISTE'));
	{$ENDIF}
	
	If PGBundleHierarchie Then 
	Begin
		If (Not PGDroitMultiForm) Then
		Begin
			SetControlText('XX_WHEREPREDEF', DossiersAInterroger('',V_PGI.NoDossier,'PSI'));
			SetControlProperty('PSI_INTERIMAIRE', 'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PSI',False,True));
		End
		Else
        Begin
			Edit := THEdit(GetControl('PSI_INTERIMAIRE'));
			If Edit <> Nil Then Edit.OnElipsisClick := ClickSalarie;
			SetControlText('XX_WHEREPREDEF', DossiersAInterroger('','','PSI'));
			SetControlProperty('PSI_INTERIMAIRE', 'Plus', DossiersAInterroger('','','PSI',False,True));
        End;
	End;
	
	{$IFDEF EAGLCLIENT}
	Liste.MultiSelect := true;
	{$ENDIF}
	SetControlVisible('BSelectAll',True);
	
	BValid := TToolBarButton97(GetControl('BVALID'));
	If BValid <> Nil Then BValid.OnClick := ValideSelection;
	
	if Liste  <>  NIL then Liste.OnDblClick  :=  GrilleDblClick;
		
	If TypeSaisie = 'SESSION' Then
	Begin
		Req := ReadTokenPipe(S,';');
		Ecran.Caption := 'Inscription des salariés à la session';
	End
	Else If TypeSaisie = 'PREVISIONNEL' Then
	Begin
		Ecran.Caption := 'Inscription des salariés à la formation';
		//SetControlVisible('CKDEJAINSCRITS', False);
		Req := ReadTokenPipe(S,';');
	End;
	UpdateCaption(Ecran);
end ;

Procedure TOF_PGMULSALINT.ValideSelection (Sender : TObject);
var i : Integer;
Salarie : String;
Begin
	if ((Liste.nbSelected) > 0) and (not Liste.AllSelected) then
	begin
		for i := 0 to Liste.NbSelected - 1 do
		begin
			{$IFDEF EAGLCLIENT}
			if (TFMul(Ecran).bSelectAll.Down) then TFMul(Ecran).Fetchlestous;
			{$ENDIF}
			Liste.GotoLeBOOKMARK(i);
			{$IFDEF EAGLCLIENT}
			TFmul(Ecran).Q.TQ.Seek(Liste.Row - 1);
			{$ENDIF}
			If Pfx = 'PSI' Then //PT1
				Salarie := TFmul(Ecran).Q.FindField('PSI_INTERIMAIRE').AsString
			Else
				Salarie := TFmul(Ecran).Q.FindField('PSA_SALARIE').AsString;
			If TFMul(Ecran).Retour ='' then TFMul(Ecran).Retour := Salarie
			else TFMul(Ecran).Retour := TFMul(Ecran).Retour+';'+Salarie;
		end;
	end
	else if liste.AllSelected then
	begin
		TFmul(Ecran).Q.First;
		while not TFmul(Ecran).Q.EOF do
		begin
			{$IFDEF EAGLCLIENT}
			if (TFMul(Ecran).bSelectAll.Down) then TFMul(Ecran).Fetchlestous;
			{$ENDIF}
			If Pfx = 'PSI' Then //PT1
				Salarie := TFmul(Ecran).Q.FindField('PSI_INTERIMAIRE').AsString
			Else
				Salarie := TFmul(Ecran).Q.FindField('PSA_SALARIE').AsString;
			If TFMul(Ecran).Retour ='' then TFMul(Ecran).Retour := Salarie
			else TFMul(Ecran).Retour := TFMul(Ecran).Retour+';'+Salarie;
			TFmul(Ecran).Q.Next;
		end;
	end;
	TFMul(Ecran).bSelectAll.Down := False;
	Liste.AllSelected := False;
	Liste.ClearSelected;
	TFMul(Ecran).BAnnulerClick(TFMul(Ecran).BAnnuler);
End;

procedure TOF_PGMULSALINT.GrilleDBLClick (Sender : TObject);
var Salarie : String;
begin
	{$IFDEF EAGLCLIENT}
	TFmul(Ecran).Q.TQ.Seek(Liste.Row-1) ;
	{$ENDIF}
	If Pfx = 'PSI' Then //PT1
		Salarie := TFmul(Ecran).Q.FindField('PSI_INTERIMAIRE').AsString
	Else
		Salarie := TFmul(Ecran).Q.FindField('PSA_SALARIE').AsString;

	TFMul(Ecran).Retour := Salarie;
	
	TFMul(Ecran).BAnnulerClick(TFMul(Ecran).BAnnuler);
end;

procedure TOF_PGMULSALINT.ClickSalarie (Sender : TObject);
Begin
	ElipsisSalarieMultidos (Sender);
End;

Initialization
  registerclasses ( [ TOF_PGMULSALINT ] ) ;

End.
