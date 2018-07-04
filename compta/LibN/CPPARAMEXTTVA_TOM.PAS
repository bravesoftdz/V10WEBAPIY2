{***********UNITE*************************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 16/01/2007
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : PARAMEXTTVA (PARAMEXTTVA)
Mots clefs ... : TOM;PARAMEXTTVA
*****************************************************************}
Unit CPPARAMEXTTVA_TOM ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes,
     db,
     CPHISTOMODIFTVA_TOF,
     CPVISUDECLATVA_TOF,
{$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     FichList,
     HDB,
     Fe_Main,  //AGLLanceFiche
{$else}
     MaineAGL, //AGLLanceFiche
     eFichList,
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     HTB97,
     UTOM,
     UTob;

  procedure LanceDeclarationTVA (TypeDeclaration : string); // SDA le 12/12/2007 ajout paramètre

Type
  TOM_PARAMEXTTVA = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
  public
    procedure BCorrectionClick 		 (Sender : TObject);
    procedure BTraitementClick 		 (Sender : TObject);
    procedure ModifEcran				 ;
    procedure TypeChange 		 		 (Sender : TObject);
    { FQ 19777 BVE 12/03/2007 }
    procedure AnneChange             (Sender : TObject);
    procedure TrimestreChange        (Sender : TObject);
    { END FQ 19777 BVE 12/03/2007 }
  private
    NumDeclaration : integer;
    Modifiable     : boolean;
    HaveModifs     : boolean;
    TypeDecl       : string; // SDA le 12/12/2007
    procedure InitEvenements; 
    function  GetNumDeclaration : integer;
end ;

Implementation

uses uLibExportTVA;

procedure LanceDeclarationTVA (TypeDeclaration : string); // SDA le 12/12/2007 ajout paramètre
var
  Arguments: string;
begin
  //SDA le 19/12/2007
  Arguments := 'TYPEDECLA=' + TypeDeclaration;
  //AGLLanceFiche('CP','CPPARAMEXTTVA','','','');
  AGLLanceFiche('CP','CPPARAMEXTTVA','','',Arguments);
  //Fin SDA le 19/12/2007
end;

procedure TOM_PARAMEXTTVA.OnArgument ( S: String ) ;
var
  x: integer;
  critere: string;
  Arg, Val: string;
  stRange: string;
begin
  Inherited ;
  // SDA le 12/12/2007
  // Gestion des arguments
  repeat
    Critere := uppercase(Trim(ReadTokenSt(S)));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        Arg := copy(Critere, 1, x - 1);
        Val := copy(Critere, x + 1, length(Critere));
        if Arg = 'TYPEDECLA' then TypeDecl := Val;
      end;
    end;
  until Critere = '';
  //Fin SDA le 12/12/2007
  InitEvenements;
  //SDA le 12/12/2007
  //Mise à jour du caption écran selon choix déclaration TVA
  if TypeDecl = '723' then  // Relevé trimestriel
  begin
    TFFicheListe(ecran).caption := 'Relevé trimestriel des livraisons intracommunautaires';
    SetField ('CPT_TYPE', TypeDecl);
  end
  else
    if TypeDecl = '725' then // Liste annuelle
    begin
      TFFicheListe(ecran).caption := 'Liste annuelle des clients assujettis';
      SetField ('CPT_TYPE', TypeDecl);
    end;
  UpdateCaption(TFFicheListe(ecran));
  //Filtre sur relevé trimestriel ou liste annuelle selon choix
  stRange := 'CPT_TYPE=''' + TypeDecl + '''';
  if stRange <> '' then
    TFFicheListe(Ecran).SetNewRange('', stRange);

end ;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 16/01/2007
Modifié le ... :   /  /
Description .. : Procedure permettant d'association des
Suite ........ : fonctions/procedures aux evenements de la fiche.
Mots clefs ... :
*****************************************************************}
procedure TOM_PARAMEXTTVA.InitEvenements;
begin
  TToolbarButton97(GetControl('BCORRECTION' )).OnClick 		:= BCorrectionClick;
  TToolbarButton97(GetControl('BTRAITEMENT' )).OnClick 		:= BTraitementClick;
{$IFNDEF EAGLCLIENT}
  THDBValComboBox(GetControl('CPT_TYPE'	  )).OnChange		:= TypeChange;
  THDBSpinEdit(GetControl('CPT_ANNEE'       )).OnChange     := AnneChange;
  THDBSpinEdit(GetControl('CPT_TRIMESTRE'   )).OnChange     := TrimestreChange;
  //SDA le 19/12/2007 version belge
  THDBEdit (GetControl('CPT_DEBPERIODE')).OnEnter := AnneChange;
  THDBEdit (GetControl('CPT_FINPERIODE')).OnEnter := AnneChange;
  //Fin SDA le 19/12/2007
{$ELSE}
  THValComboBox(GetControl('CPT_TYPE'	  	  )).OnChange		:= TypeChange;
  THSpinEdit(GetControl('CPT_ANNEE'         )).OnChange     := AnneChange;
  THSpinEdit(GetControl('CPT_TRIMESTRE'     )).OnChange     := TrimestreChange;
  //SDA le 19/12/2007 version belge
  THEdit (GetControl('CPT_DEBPERIODE')).OnEnter := AnneChange;
  THEdit (GetControl('CPT_FINPERIODE')).OnEnter := AnneChange;
  //Fin SDA le 19/12/2007
{$ENDIF}
  //TFFicheListe(ecran).StaDataChange                         := ModifEcran;
end;

procedure TOM_PARAMEXTTVA.TypeChange;
begin
{FQ 19776 BVE 12032007}

{$IFNDEF EAGLCLIENT}
	if (THDBValComboBox(GetControl('CPT_TYPE')).Value = '723') then
   begin
      if Modifiable then
      begin
   	   SetControlEnabled('CPT_TRIMESTRE',true);
       if (DS.State<>dsBrowse) then // SDA le 21/01/2008
         SetField('CPT_TRIMESTRE',GetControlText('CPT_TRIMESTRE'));
      end;
      SetControlEnabled('BCORRECTION',true);
   end
   else if THDBValComboBox(GetControl('CPT_TYPE')).Value = '725' then
   begin
   	SetControlEnabled('CPT_TRIMESTRE',false);
      SetControlEnabled('BCORRECTION',false);
   end;
{$ELSE}
	if (THValComboBox(GetControl('CPT_TYPE')).Value = '723') then
   begin
      if Modifiable then
      begin
   	   SetControlEnabled('CPT_TRIMESTRE',true);
       if (DS.State<>dsBrowse) then // SDA le 21/01/2008
         SetField('CPT_TRIMESTRE',GetControlText('CPT_TRIMESTRE'));
      end;
      SetControlEnabled('BCORRECTION',true);
   end
   else if THValComboBox(GetControl('CPT_TYPE')).Value = '725' then
   begin
   	SetControlEnabled('CPT_TRIMESTRE',false); 
      SetControlEnabled('BCORRECTION',false);
   end;
{$ENDIF}
{END FQ 19776 BVE 12032007}
   AnneChange(nil);
end;


procedure TOM_PARAMEXTTVA.AnneChange             (Sender : TObject);
var
{$IFNDEF EAGLCLIENT}
   Annee            : THDBSpinEdit; 
   Trimestre        : THDBSpinEdit;
   DebPeriode       : THDBEdit;   
   FinPeriode       : THDBEdit;
{$ELSE}                 
   Annee            : THSpinEdit; 
   Trimestre        : THSpinEdit;
   DebPeriode       : THEdit;
   FinPeriode       : THEdit;
{$ENDIF}
begin

   if not(Modifiable) then Exit;
   if (DS.State=dsBrowse) then exit;// SDA le 21/01/2008
   //if (DS.State = dsBrowse) then  exit;// SDA le 21/01/2008
{$IFNDEF EAGLCLIENT}
   Annee      := THDBSpinEdit(GetControl('CPT_ANNEE'));
   Trimestre  := THDBSpinEdit(GetControl('CPT_TRIMESTRE'));
   DebPeriode := THDBEdit(GetControl('CPT_DEBPERIODE'));
   FinPeriode := THDBEdit(GetControl('CPT_FINPERIODE'));
{$ELSE}
   Annee      := THSpinEdit(GetControl('CPT_ANNEE')); 
   Trimestre  := THSpinEdit(GetControl('CPT_TRIMESTRE'));
   DebPeriode := THEdit(GetControl('CPT_DEBPERIODE'));
   FinPeriode := THEdit(GetControl('CPT_FINPERIODE'));
{$ENDIF}
   if Annee.Value < GetMinAnneeExercice then
   begin
      Annee.Value := GetMinAnneeExercice;
      Exit;
   end
   else if  Annee.Value > GetMaxAnneeExercice then
   begin
      Annee.Value := GetMaxAnneeExercice;
      Exit;
   end;

   if GetControlEnabled('CPT_TRIMESTRE') then
   begin
      // On est sur une déclaration trimestrielle
      //SDA le 21/12/2004
      //DebPeriode.Text := DateToStr(EncodeDate(Annee.Value,(Trimestre.Value * 3 - 2),1));
      SetField ('CPT_DEBPERIODE', DateToStr(EncodeDate(Annee.Value,(Trimestre.Value * 3 - 2),1)));
      if Trimestre.Value < 4 then
         //SDA le 21/12/2007
         //FinPeriode.Text := DateToStr(EncodeDate(Annee.Value,(Trimestre.Value * 3 + 1),1) - 1)
         SetField ('CPT_FINPERIODE', DateToStr(EncodeDate(Annee.Value,(Trimestre.Value * 3 + 1),1) - 1))
      else
         //SDA le 21/12/2007
         //FinPeriode.Text := DateToStr(EncodeDate(Annee.Value + 1,1,1) - 1);
         SetField ('CPT_FINPERIODE', DateToStr(EncodeDate(Annee.Value + 1,1,1) - 1));
   end
   else         
   begin
      // On est sur une déclaration annuelle
      //SDA le 21/12/2007
      //DebPeriode.Text := DateToStr(EncodeDate(Annee.Value,1,1));
      SetField ('CPT_DEBPERIODE', DateToStr(EncodeDate(Annee.Value,1,1)));
      //SDA le 21/12/2007
      //FinPeriode.Text := DateToStr(EncodeDate(Annee.Value+1,1,1) - 1);
      SetField ('CPT_FINPERIODE', DateToStr(EncodeDate(Annee.Value+1,1,1) - 1));
   end;

end;

procedure TOM_PARAMEXTTVA.TrimestreChange        (Sender : TObject);
var
{$IFNDEF EAGLCLIENT}
   Annee            : THDBSpinEdit; 
   Trimestre        : THDBSpinEdit;
   DebPeriode       : THDBEdit;   
   FinPeriode       : THDBEdit;
{$ELSE}                 
   Annee            : THSpinEdit; 
   Trimestre        : THSpinEdit;
   DebPeriode       : THEdit;
   FinPeriode       : THEdit;
{$ENDIF}
begin

   if not(Modifiable) then Exit;
   //if (DS.State = dsBrowse) then  exit;// SDA le 21/01/2008
{$IFNDEF EAGLCLIENT}
   Annee      := THDBSpinEdit(GetControl('CPT_ANNEE'));
   Trimestre  := THDBSpinEdit(GetControl('CPT_TRIMESTRE'));
   DebPeriode := THDBEdit(GetControl('CPT_DEBPERIODE'));
   FinPeriode := THDBEdit(GetControl('CPT_FINPERIODE'));
{$ELSE}
   Annee      := THSpinEdit(GetControl('CPT_ANNEE')); 
   Trimestre  := THSpinEdit(GetControl('CPT_TRIMESTRE'));
   DebPeriode := THEdit(GetControl('CPT_DEBPERIODE'));
   FinPeriode := THEdit(GetControl('CPT_FINPERIODE'));
{$ENDIF}
   if (Trimestre.Value < 1) then  
   begin
      Trimestre.Value := 1;
      Exit;
   end
   else if (Trimestre.Value > 4) then
   begin
      Trimestre.Value := 4;
      Exit;
   end;
   // On est sur une déclaration trimestrielle
   DebPeriode.Text := DateToStr(EncodeDate(Annee.Value,(Trimestre.Value * 3 - 2),1));
   if Trimestre.Value < 4 then
      FinPeriode.Text := DateToStr(EncodeDate(Annee.Value,(Trimestre.Value * 3 + 1),1) - 1)
   else
      FinPeriode.Text := DateToStr(EncodeDate(Annee.Value + 1,1,1) - 1);
end;

procedure TOM_PARAMEXTTVA.ModifEcran;
begin
   SetControlEnabled('BTRAITEMENT',false);
end;

procedure TOM_PARAMEXTTVA.OnNewRecord ;
var
{$IFNDEF EAGLCLIENT}
  Code   : THDBEdit;
  Annee  : THDBSpinEdit;
{$ELSE}
  Code : THEdit;
  Annee  : THSpinEdit;
{$ENDIF}

begin

  Inherited ;

{$IFNDEF EAGLCLIENT}
  // Désactivation du Code des modifications et renseignement automatiquue
  Code := THDBEdit(GetControl('CPT_CODE'));
  Code.Enabled := false;
  Code.Field.Value := IntToStr(GetCodeMax('PARAMEXTTVA')+1);

  // On positionne le spinedit année avec les valeurs recuperées dans la base.
  Annee := THDBSpinEdit(GetControl('CPT_ANNEE'));
  Annee.MaxValue := GetMaxAnneeExercice;
  Annee.MinValue := GetMinAnneeExercice;
  Annee.Value := GetMinAnneeExercice;
  SetField('CPT_ANNEE',GetControlText('CPT_ANNEE'));

{$ELSE}
  // Désactivation du Code des modifications et renseignement automatiquue
  Code := THEdit(GetControl('CPT_CODE'));
  Code.Enabled := false;
  Code.Text := IntToStr(GetCodeMax('PARAMEXTTVA')+1);
  SetField('CPT_CODE', GetControlText('CPT_CODE')); //SDA le 21/12/2007

  // On positionne le spinedit année avec les valeurs recuperées dans la base.
  Annee := THSpinEdit(GetControl('CPT_ANNEE'));
  Annee.MaxValue := GetMaxAnneeExercice;
  Annee.MinValue := GetMinAnneeExercice;
  Annee.Value := GetMinAnneeExercice;
  SetField('CPT_ANNEE',GetControlText('CPT_ANNEE'));

{$ENDIF}

  //SDA le 12/12/2007
  if TypeDecl = '723' then
    SetField ('CPT_TYPE', TypeDecl)
  else
    if TypeDecl = '725' then
      SetField ('CPT_TYPE',TypeDecl);
  //Fin SDA le 12/12/2007

  //SDA le 21/12/2007 gestion nouvelle zone IDDECLA
  THEdit(GetControl('IDDECLA')).Text := GetControlText('CPT_CODE');
  //Fin SDA le 21/12/2007

end ;

procedure TOM_PARAMEXTTVA.OnDeleteRecord ;
begin
  Inherited ;
  // Il faut aussi supprimer l'ensemble des modifications de la declaration
  DeleteModifsDeclaration(NumDeclaration);
end ;

procedure TOM_PARAMEXTTVA.OnUpdateRecord ;
var
{$IFNDEF EAGLCLIENT}
  CodeDecla   : THDBEdit;
{$ELSE}
  CodeDecla : THEdit;
{$ENDIF}
begin
  Inherited ;
  //SDA le 21/12/2007
  if (DS.State in [dsInsert]) then
  begin
    {$IFNDEF EAGLCLIENT}
    // Désactivation du Code des modifications et renseignement automatiquue
    CodeDecla := THDBEdit(GetControl('CPT_CODE'));
    CodeDecla.Enabled := false;
    CodeDecla.Field.Value := IntToStr(GetCodeMax('PARAMEXTTVA')+1);
    SetField('CPT_CODE', GetControlText('CPT_CODE')); //SDA le 21/12/2007
    {$ELSE}
    // Désactivation du Code des modifications et renseignement automatiquue
    CodeDecla := THEdit(GetControl('CPT_CODE'));
    CodeDecla.Enabled := false;
    CodeDecla.Text := IntToStr(GetCodeMax('PARAMEXTTVA')+1);
    SetField('CPT_CODE', GetControlText('CPT_CODE')); //SDA le 21/12/2007
    {$ENDIF}
    THEdit(GetControl('IDDECLA')).Text := GetField('CPT_CODE');
    SetFocusControl('CPT_LIBELLE'); //SDA le 21/12/2007
  end;
  //Fin SDA le 21/12/2007

end ;

procedure TOM_PARAMEXTTVA.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_PARAMEXTTVA.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_PARAMEXTTVA.OnLoadRecord ;
begin

  Inherited ;

  NumDeclaration := GetNumDeclaration;
  Modifiable := IsModifEditable(NumDeclaration);
  HaveModifs := DeclarationHaveModif(NumDeclaration);

  // On grise ou non les zones
{ FQ 19777 BVE 12/03/2007
  if (DS.State = dsInsert) then
    SetControlEnabled('BCORRECTION',false)
  else
    SetControlEnabled('BCORRECTION',Modifiable or HaveModifs);
	if (DS.State in [dsInsert, dsEdit]) then
      SetControlEnabled('BTRAITEMENT',false)
	else
      SetControlEnabled('BTRAITEMENT',true);
END FQ 19777 BVE 12/03/2007 }
  SetControlEnabled('BCORRECTION',Modifiable or HaveModifs);
  SetControlEnabled('BTRAITEMENT',true);
  SetControlEnabled('CPT_CODE',False);
  SetControlEnabled('CPT_LIBELLE',Modifiable);
  SetControlEnabled('CPT_ABREGE',Modifiable);
  //SDA le 21/12/2007 SetControlEnabled('CPT_TYPE',Modifiable);
  SetControlEnabled('CPT_ANNEE',Modifiable);
  SetControlEnabled('CPT_TRIMESTRE',Modifiable);
  SetControlEnabled('CPT_DEBPERIODE',Modifiable);
  SetControlEnabled('CPT_FINPERIODE',Modifiable);
  SetControlEnabled('CPT_NATURE',Modifiable);
  SetControlEnabled('CPT_COLLECTIF',Modifiable);
  SetControlEnabled('CPT_REGIMETVA',Modifiable);
  SetControlEnabled('BDELETE',Modifiable);

  // Affichage ou non des libelles d'informations.
  SetControlVisible('VALIDEE',not(Modifiable));
  SetControlVisible('MODIFS', HaveModifs);

  TypeChange(nil);

  //SDA le 19/12/2007 version belge
  SetFocusControl('CPT_LIBELLE'); // positionnement focus sur libellé
  //Fin SDA le 19/12/2007

  //SDA le 21/12/2007
  //if (DS.State <> dsBrowse) then // SDA le 21/01/2008
  //begin
    if  GetControlText ('CPT_CODE') <> '' then
      THEdit(GetControl('IDDECLA')).Text := GetControlText ('CPT_CODE');
  //end;
  //Fin SDA le 21/12/2007

end ;

procedure TOM_PARAMEXTTVA.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_PARAMEXTTVA.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_PARAMEXTTVA.OnCancelRecord ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 16/01/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
procedure TOM_PARAMEXTTVA.BTraitementClick;
begin
   // On sauvergarde la fiche au cas ou il y ai eu une modif sans enregistrement.
   TFFicheListe(ecran).BValiderClick(ecran);
{ FQ 19775 BVE 08/03/2007
   if (Modifiable) then
      Traitement := PGIASK('Etes-vous sur de vouloir valider cette déclaration ?')
   else
      Traitement := PGIASK('Voulez-vous rééditer cette déclaration ?');
}
   LanceVisuDeclaTVA('CODECPT=' + IntToStr(GetNumDeclaration));
   // On relance l'affichage afin de prendre en compte les modifs.
   RefreshDB;
   // On force l'enabled du bouton delete.
  	SetControlEnabled('BDELETE',false);

end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 16/01/2007
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOM_PARAMEXTTVA.BCorrectionClick;
begin
  if not(HaveModifs) and not(Modifiable) then
  begin
     PGIInfo('Il n''y a pas de corrections associées à cette déclaration');
     Exit;
  end;
{ FQ 19777 BVE 12/03/2007 }
  // On force l'enregistrement
  TFFicheListe(ecran).BValiderClick(ecran);
{END FQ 19777 BVE 12/03/2007}
  LanceModificationTVA('CODECPT=' + IntToStr(NumDeclaration));
  refreshDB;//SDA le 21/01/2008
end;

function TOM_PARAMEXTTVA.GetNumDeclaration : integer;
begin
  result := 0; // SDA le 21/12/2007
{$IFNDEF EAGLCLIENT}
  if THDBEdit(GetControl('CPT_CODE')).Text <> '' then // SDA le 21/12/2007
    result := StrToInt(THDBEdit(GetControl('CPT_CODE')).Text);
{$ELSE}
  if THEdit(GetControl('CPT_CODE')).Text <> '' then //SDA le 21/12/2007
    result := StrToInt(THEdit(GetControl('CPT_CODE')).Text);
{$ENDIF}
end;


Initialization
  registerclasses([TOM_PARAMEXTTVA]);
end.
