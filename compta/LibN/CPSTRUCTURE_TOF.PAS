{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 04/04/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPSTRUCTURE ()
Mots clefs ... : TOF;CPSTRUCTURE
*****************************************************************}
Unit CPSTRUCTURE_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Windows,        // VK_
     Vierge,         // TFVierge
{$IFDEF EAGLCLIENT}
     MaineAGL,       // AGLLanceFiche
     UtileAGL,       // TNavigateBtn, PrintDBGrid
{$ELSE}
     db,
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
     FE_Main,        // AGLLanceFiche
     DBCtrls,        // TNavigateBtn
     { BVE 02.05.07 Section,        // CodeStructure}
     PrintDBG,       // PrintDBGrid
{$ENDIF}
     HStatus,        // InitMove, MoveCur, FiniMove
{$IFNDEF CCMP}
  {$IFNDEF GCGC}
     GeSecAna,       // GenereSectionsAna
  {$ENDIF}
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,         // ReadTokenSt
     HEnt1,
     HMsgBox,
     HTB97,          // TToolBarButton97
     HPanel,         // THPanel
     utilPGI,        // _Blocage
     Ent1,           // VH
     UTOB,
     UTOF ;

Procedure ParamPlanAnal(FAxe : String);

Type
  TOF_CPSTRUCTURE = Class (TOF)
    FListe1   : THGrid;
    FListe2   : THGrid;
    PAxe      : TPageControl;

    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure InitMsgBox;
    procedure InitListe;
    procedure TPlanOpen(bCharge : Boolean = False);
    procedure TSSPlanOpen;
    function  TPlanModified : Boolean;
    function  TSSPlanModified : Boolean;

    procedure BDefaireClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure BValider1Click(Sender: TObject);
    procedure BImprimer1Click(Sender: TObject);

    procedure BDefaire1Click(Sender: TObject);
    procedure BInsert1Click(Sender: TObject);
    procedure BDelete1Click(Sender: TObject);
    procedure BAutomateClick(Sender: TObject);
    procedure BMultiSectClick(Sender: TObject);
    procedure BSimpSectClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    
    procedure PAxeChange(Sender: TObject);

    function  Bouge1(Button: TNavigateBtn) : boolean;
    function  OnSauve1 : boolean ;
    function  EnregOK1 : boolean ;
    function  NewEnreg1Possible : Boolean;
    procedure NewEnreg1 ;
    function  VerifiSiExiste1 : Boolean ;

    function  Bouge2(Button: TNavigateBtn) : boolean;
    function  OnSauve2 : boolean ;
    procedure NewEnreg2 ;
    function  EnregOK2 : boolean ;
    function  VerifiSiExiste2 : Boolean ;

    procedure FListe1DblClick(Sender: TObject);
    procedure FListe1RowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure FListe1RowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure FListe1KeyPress(Sender: TObject; var Key: Char);
    procedure FListe1KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);

    procedure FListe2RowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure FListe2KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  private
    MsgBox   : THMsgBox;
    MsgBox1  : THMsgBox;
    MsgBox2  : THMsgBox;
    FAxe     : String;
    FAvertir : Boolean;
    FClosing : Boolean;
    FLastRep : Integer;
    MemoLg   : Integer;
    FLgSect  : Integer;
    LgDeb    : Integer;
    TPlan    : Tob;
    TSSPlan  : Tob;

    procedure LongueurSection;
    function  SommeLongeur : Integer;
    function  ValideLg : Boolean;
    function  ChercheEnreg2 : Boolean;
    procedure DetruitContenuPlan;
    procedure SwapLaValeur;
    function  ChercheUnTrou : Byte;
    function  TrouDansEnchainement(Var NumMess : Byte) : Boolean;
    function  BourreLeCompte : Boolean;
    function  MauvaiseLg : Boolean;
    function  OnlyOne  : Boolean;
    procedure ErgoS3S5;
    procedure GenereCodeContenuPlan;
    procedure ActiveBouton ;
{$IFDEF CCS3}
    function  NbSousPlan : Integer ;
{$ENDIF}
end ;

Implementation

uses        
  {$IFDEF MODENT1}
    CPProcMetier,
  {$ENDIF}
    CPSection_TOM,
    {$IFDEF eAGLCLIENT}
    MenuOLX
    {$ELSE}
    MenuOLG
    {$ENDIF eAGLCLIENT}
    ;
var OldIndex : Integer ;

Procedure ParamPlanAnal(FAxe : String);
begin
  if _Blocage(['nrCloture'],False,'nrAucun') then Exit ;
  AGLLanceFiche('CP','CPSTRUCTURE','','',FAxe);
end;

procedure TOF_CPSTRUCTURE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPSTRUCTURE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPSTRUCTURE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPSTRUCTURE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPSTRUCTURE.OnArgument (S : String ) ;
begin
  Inherited ;
  // Création des contrôles
  MsgBox  := THMsgBox.create(FMenuG);
  MsgBox1 := THMsgBox.create(FMenuG);
  MsgBox2 := THMsgBox.create(FMenuG);

  InitMsgBox;
  Ecran.HelpContext := 1445000;

  // Récupère les arguments
  FAxe := ReadTokenSt(S);

  // Evénements des contrôles
  TToolBarButton97(GetControl('BDEFAIRE',True)).OnClick := BDefaireClick;
  TToolBarButton97(GetControl('BINSERT',True)).OnClick := BInsertClick;
  TToolBarButton97(GetControl('BDELETE',True)).OnClick := BDeleteClick;
  TToolBarButton97(GetControl('BVALIDER1',True)).OnClick := BValider1Click;
  TToolBarButton97(GetControl('BIMPRIMER1',True)).OnClick := BImprimer1Click;

  TToolBarButton97(GetControl('BDEFAIRE1',True)).OnClick := BDefaire1Click;
  TToolBarButton97(GetControl('BINSERT1',True)).OnClick := BInsert1Click;
  TToolBarButton97(GetControl('BDELETE1',True)).OnClick := BDelete1Click;
  TToolBarButton97(GetControl('BAUTOMATE',True)).OnClick := BAutomateClick;
  TToolBarButton97(GetControl('BMULTISECT',True)).OnClick := BMultiSectClick;
  TToolBarButton97(GetControl('BSIMPSECT',True)).OnClick := BSimpSectClick;
  TToolBarButton97(GetControl('BIMPRIMER',True)).OnClick := BImprimerClick;
  TToolBarButton97(GetControl('BVALIDER',True)).OnClick := BValiderClick;
  TToolBarButton97(GetControl('HELPBTN',True)).OnClick := BAideClick;

  PAxe := TPageControl(GetControl('PAXE',True));
  
  PAxe.OnChange := PAxeChange;

  // Propriétés des grilles
  FListe1 := THGrid(GetControl('FLISTE1',True));
  FListe2 := THGrid(GetControl('FLISTE2',True));

  FListe1.OnDblClick := FListe1DblClick;
  FListe1.OnRowEnter := FListe1RowEnter;
  FListe1.OnRowExit  := FListe1RowExit;
  FListe1.OnKeyPress := FListe1KeyPress;
  FListe1.OnKeyDown := FListe1KeyDown;

  FListe2.OnRowExit  := FListe2RowExit;
  FListe2.OnKeyDown := FListe2KeyDown;

  InitListe;

  // Initialisation
  FAvertir := False;
  FClosing := False;
  FLastRep := 0;
  if (FAxe = '') then FAxe:='A1'
                 else PAxe.ActivePageIndex := StrToInt(Copy(FAxe,2,1))-1;

  LongueurSection;
  ErgoS3S5;
  TFVierge(Ecran).HMTrad.ResizeGridColumns(Fliste1);
  TFVierge(Ecran).HMTrad.ResizeGridColumns(Fliste2);

  TPlan := Tob.Create('STRUCRSE',nil,-1);
  TPlanOpen(True);

  TSSPlan := Tob.Create('SSSTRUCR',nil,-1);
  TSSPlanOpen;
end ;

procedure TOF_CPSTRUCTURE.OnClose ;
var
  i : Byte ;
  CanClose : Boolean;
begin
  Inherited ;
  FClosing := True;

  CanClose := OnSauve1;
  if not CanClose then begin LastError := -1; Exit; end;

  CanClose := OnSauve2;
  if not CanClose then begin LastError := -1; Exit; end;

  if TrouDansEnchainement(i) then begin
    Canclose := False ;
    MsgBox1.Execute(i,'','') ;
  end;
  // YMO 28/09/05 FQ 16523 (point 4) Pb de râfraichissement du cache
  AvertirCacheServer('STRUCRSE');
  AvertirCacheServer('SSSTRUCR');
  ChargeSousPlanAxe;
  if not CanClose then begin LastError := -1; Exit; end;

  // S'il y a eu des changements : Recharge les données
  if FAvertir then begin
    ChargeStructureUnique ;
    AvertirTable('ttStrucrse1') ;
    AvertirTable('ttStrucrse2') ;
    AvertirTable('ttStrucrse3') ;
    AvertirTable('ttStrucrse4') ;
    AvertirTable('ttStrucrse5') ;
  end;
  MsgBox.Free;
  MsgBox1.Free;
  MsgBox2.Free;
  TPlan.Free;
  TSSPlan.Free;
end ;

procedure TOF_CPSTRUCTURE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPSTRUCTURE.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPSTRUCTURE.InitMsgBox;
begin
  MsgBox.Mess.Add('0;'+Ecran.Caption+';Cet axe analytique n''est pas structuré.;W;O;O;O;');
  MsgBox.Mess.Add('L''enregistrement est inaccessible');
{UTILISE}  MsgBox.Mess.Add('Contenu du plan : ');
{UTILISE}  MsgBox.Mess.Add('Caractères');
{UTILISE}  MsgBox.Mess.Add('Plans de sous-sections : ');

MsgBox1.Mess.Add('0;'+Ecran.Caption+';Voulez-vous enregistrer les modifications ?;Q;YNC;Y;C;');
MsgBox1.Mess.Add('1;'+Ecran.Caption+';Confirmez-vous la suppression de l''enregistrement ?;Q;YNC;N;C;');
MsgBox1.Mess.Add('2;'+Ecran.Caption+';Vous devez renseigner un code.;W;O;O;O;');
MsgBox1.Mess.Add('3;'+Ecran.Caption+';Vous devez renseigner un libellé.;W;O;O;O;');
MsgBox1.Mess.Add('4;'+Ecran.Caption+';Le code que vous avez saisi existe déjà. Vous devez le modifier.;W;O;O;O;');
MsgBox1.Mess.Add('5;'+Ecran.Caption+';Vous avez déjà atteint la longueur maximum des sections de cet axe.;W;O;O;O;');
MsgBox1.Mess.Add('6;'+Ecran.Caption+';La longueur que vous avez renseignée est trop grande.;W;O;O;O;');
MsgBox1.Mess.Add('7;'+Ecran.Caption+';La modification de la longueur va entraîner la destruction du contenu des plans de sous-sections. Désirez-vous continuez ?;Q;YNC;N;C;');
MsgBox1.Mess.Add('8;'+Ecran.Caption+';Votre axe est mal structuré : il n''y a pas de continuité dans les enchaînements début longueur.;W;O;O;O;');
MsgBox1.Mess.Add('9;'+Ecran.Caption+';Une des longueurs que vous avez renseignées n''est pas valide.;W;O;O;O;');
MsgBox1.Mess.Add('10;'+Ecran.Caption+';Vous devez renseigner un code.;W;O;O;O;');
MsgBox1.Mess.Add('11;'+Ecran.Caption+';Vous ne pouvez pas renseigner un longueur égale à 0.;W;O;O;O;');
MsgBox1.Mess.Add('12;'+Ecran.Caption+';Vous ne pouvez pas renseigner plus de 3 sous-plans.;W;O;O;O;');
MsgBox1.Mess.Add('13;'+Ecran.Caption+';Vous devez renseigner une longueur.;W;O;O;O;');

MsgBox2.Mess.Add('0;'+Ecran.Caption+';Voulez-vous enregistrer les modifications ?;Q;YNC;Y;C;');
MsgBox2.Mess.Add('1;'+Ecran.Caption+';Confirmez-vous la suppression de l''enregistrement ?;Q;YNC;N;C;');
MsgBox2.Mess.Add('2;'+Ecran.Caption+';Vous devez renseigner un code.;W;O;O;O;');
MsgBox2.Mess.Add('3;'+Ecran.Caption+';Vous devez renseigner un libellé.;W;O;O;O;');
MsgBox2.Mess.Add('4;'+Ecran.Caption+';Le code que vous avez saisi existe déjà. Vous devez le modifier.;W;O;O;O;');
MsgBox2.Mess.Add('5;'+Ecran.Caption+';Désirez-vous générer les codes de sous-sections à partir des sections analytiques déjà définies ?;Q;YNC;Y;C;');
MsgBox2.Mess.Add('6;'+Ecran.Caption+';Le code choisi ne respecte pas la longueur de cette sous-section.;W;O;O;O;');
MsgBox2.Mess.Add('Génération des codes');
MsgBox2.Mess.Add('8;'+Ecran.Caption+';Désirez-vous générer toutes les combinaisons des sections analytiques incluant le sous plan sélectionné ?;Q;ON;N;N;');
end;

procedure TOF_CPSTRUCTURE.BDefaireClick(Sender: TObject);
begin
  Bouge1(nbCancel);
end;

procedure TOF_CPSTRUCTURE.BInsertClick(Sender: TObject);
begin
  Bouge1(nbInsert);
end;

procedure TOF_CPSTRUCTURE.BDeleteClick(Sender: TObject);
begin
  Bouge1(nbDelete);
end;

procedure TOF_CPSTRUCTURE.BValider1Click(Sender: TObject);
begin
  Bouge1(nbPost);
end;

// Impression des plans de sous-sections
procedure TOF_CPSTRUCTURE.BImprimer1Click(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
  LanceEtatTob('E','CST','SAN',TPlan,True,false,false,nil,'',Ecran.Caption+' '+PAxe.ActivePage.Caption,False,0,'',0);
//  PrintDBGrid(Ecran.Caption+' '+PAxe.ActivePage.Caption,FListe1,'');
{$ELSE}
  PrintDBGrid(FListe1, Nil, Ecran.Caption+' '+PAxe.ActivePage.Caption, '');
{$ENDIF}
end;

procedure TOF_CPSTRUCTURE.BDefaire1Click(Sender: TObject);
begin
  Bouge2(nbCancel);
end;

procedure TOF_CPSTRUCTURE.BInsert1Click(Sender: TObject);
begin
  Bouge2(nbInsert);
end;

procedure TOF_CPSTRUCTURE.BDelete1Click(Sender: TObject);
begin
  Bouge2(nbDelete);
end;

procedure TOF_CPSTRUCTURE.BAutomateClick(Sender: TObject);
begin
  if MsgBox2.Execute(5,'','')=mrYes then begin // Désirez-vous générer les codes de sous-sections à partir des sections analytiques déjà définies ?
    GenereCodeContenuPlan;
    TSSPlanOpen;
  end;
end;

procedure TOF_CPSTRUCTURE.BMultiSectClick(Sender: TObject);
begin
{$IFNDEF CCMP}
  {$IFNDEF GCGC}
    if MsgBox2.Execute(8,'','')=mrOk then // Désirez-vous générer toutes les combinaisons des sections analytiques incluant le sous plan sélectionné ?
    GenereSectionsAna(FAxe, FListe1.Cells[2,FListe1.Row], FListe2.Cells[3,FListe2.Row], FListe2.Cells[4,FListe2.Row], False);
  {$ENDIF}
{$ENDIF}
end;

procedure TOF_CPSTRUCTURE.BSimpSectClick(Sender: TObject);
begin
{$IFNDEF CCMP}
  {$IFNDEF GCGC}
    GenereSectionsAna(FAxe,'','','',False);
  {$ENDIF}
{$ENDIF}
end;

// Impression du contenu des plans de sous-sections
procedure TOF_CPSTRUCTURE.BImprimerClick(Sender: TObject);
var
  sz : String ;
begin
  if (TPlan.Detail.Count = 0) then sz := ''
                              else sz := ' - Sous-section : '+FListe1.Cells[2,FListe1.Row];
{$IFDEF EAGLCLIENT}
  LanceEtatTob('E','CST','SSA',TSSPlan,True,false,False,nil,'',Ecran.Caption+' '+PAxe.ActivePage.Caption+sz,False,0,'',0);
//  PrintDBGrid(Ecran.Caption+' '+PAxe.ActivePage.Caption+sz,FListe2,'');
{$ELSE}
  PrintDBGrid(FListe2, Nil, Ecran.Caption+' '+PAxe.ActivePage.Caption+sz, '');
{$ENDIF}
end;

procedure TOF_CPSTRUCTURE.BValiderClick(Sender: TObject);
begin
  Bouge2(nbPost);
end;

procedure TOF_CPSTRUCTURE.BAideClick(Sender: TObject);
begin
  CallHelpTopic(Ecran);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann MORENO
Créé le ...... : 28/09/2005
Modifié le ... : 28/09/2005
Description .. : Suppression de OnPAxeChanging : Pb de màj des onglets
Suite ........ : par click, et transfert des fonctionnalités dans
Suite ........ : OnPAxeChange
Mots clefs ... : YMO Change
*****************************************************************}
procedure TOF_CPSTRUCTURE.PAxeChange(Sender: TObject);
var
  NewAxe : String ;
  i : Byte ;
begin
  NewAxe := 'A'+IntToStr(PAxe.ActivePageIndex+1);

  if (FListe1.ColEditables[2]) or (FListe2.ColEditables[3]) then
  begin
    Paxe.ActivePageIndex:=OldIndex;
    Exit;
  end;

  if TrouDansEnchainement(i) then
  begin
    Paxe.ActivePageIndex:=OldIndex;
    MsgBox1.Execute(i,'','');
    Exit;
  end;

  if Not ValideLg then
  begin
    Paxe.ActivePageIndex:=OldIndex;
    Bouge1(nbCancel);
    Exit;
  end;

  if not (VH^.Cpta[AxeToFb(NewAxe)].Structure) then begin
    Paxe.ActivePageIndex:=OldIndex;
    MsgBox.Execute(0,'',''); // Cet axe analytique n'est pas structuré.
    Exit;
  end;


  if (VH^.Cpta[AxeToFb(NewAxe)].Structure) then begin
    FAxe := NewAxe ;
    FListe1.Row := 1;
    TPlanOpen(True);
    TSSPlanOpen;
  end;
 
  LongueurSection;

end;

procedure TOF_CPSTRUCTURE.LongueurSection;
var
  i : Byte;
begin
  FLgSect := VH^.Cpta[AxeToFb(FAxe)].Lg ;
  if FLgSect<=0 then for i := 1 to 5 do SetControlVisible('LGAXE'+IntToStr(i),False)
                else for i := 1 to 5 do SetControlVisible('LGAXE'+IntToStr(i),True);

  // xx Caractères
  SetControlText('LGAXE'+FAxe[2],IntToStr(FLgSect)+' '+MsgBox.Mess[3]);

end;

procedure TOF_CPSTRUCTURE.ErgoS3S5;
begin
  if EstSerie(S3) then begin
    PAxe.ActivePageIndex := 0;
    PAxe.Visible := False;
    SetControlVisible('BMULTISECT',False);
    SetControlVisible('BAUTOMATE',False);
  end;
end;

procedure TOF_CPSTRUCTURE.InitListe;
var
  i : Byte;
begin
  FListe1.ColTypes[4] := 'B';
  FListe1.ColFormats[4] := '2';
  FListe1.ColTypes[5] := 'R';
  FListe1.ColTypes[6] := 'R';
  FListe1.ColEditables[2] := False;

  FListe1.ColLengths[1] := -1; // SS_AXE
  FListe1.ColLengths[2] := 3;  // SS_SOUSSECTION
  FListe1.ColLengths[3] := 35; // SS_LIBELLE
  FListe1.ColLengths[4] := 1;  // SS_CONTROLE
  FListe1.ColLengths[5] := 3;  // SS_DEBUT
  FListe1.ColLengths[6] := 3;  // SS_LONGUEUR
  FListe1.ColLengths[7] := -1; // SS_SOCIETE

  FListe1.ColWidths[1] := -1;
  FListe1.ColWidths[7] := -1;

  FListe2.ColLengths[1] := -1;  // PS_AXE
  FListe2.ColLengths[2] := -1;  // PS_SOUSSECTION
  FListe2.ColLengths[3] := 3;   // PS_CODE
  FListe2.ColLengths[4] := 35;  // PS_LIBELLE
  FListe2.ColLengths[5] := -1;  // PS_ABREGE
  FListe2.ColLengths[6] := -1;  // PS_SOCIETE
  FListe2.ColLengths[7] := -1;  // PS_CORRESP
  FListe2.ColLengths[8] := -1;  // PS_TOTDEBP
  FListe2.ColLengths[9] := -1;  // PS_TOTCREP
  FListe2.ColLengths[10] := -1; // PS_TOTDEBE
  FListe2.ColLengths[11] := -1; // PS_TOTCREE
  FListe2.ColLengths[12] := -1; // PS_TOTDEBS
  FListe2.ColLengths[13] := -1; // PS_TOTCRES
  FListe2.ColLengths[14] := -1; // PS_TOTDEBANO
  FListe2.ColLengths[15] := -1; // PS_TOTCREANO
  FListe2.ColLengths[16] := -1; // PS_TOTDEBANON1
  FListe2.ColLengths[17] := -1; // PS_TOTCREANON1
  FListe2.ColLengths[18] := -1; // PS_DOMAINE

  for i := 1 to FListe2.ColCount-1 do
    if FListe2.ColLengths[i] = -1 then FListe2.ColWidths[i] := -1;
  FListe2.ColEditables[3] := False;
end;

// Charge la Liste1 en fonction de l'axe sélectionné
procedure TOF_CPSTRUCTURE.TPlanOpen(bCharge : Boolean = False);
var
  OldRow : Integer;
begin
  OldRow := FListe1.Row;
  if bCharge then TPlan.LoadDetailDB('STRUCRSE','"'+FAxe+'"','',nil,False,False);
  TPlan.PutGridDetail(FListe1,False,False,'SS_AXE;SS_SOUSSECTION;SS_LIBELLE;SS_CONTROLE;SS_DEBUT;SS_LONGUEUR;SS_SOCIETE',True);
  if (FListe1.RowCount > OldRow) then FListe1.Row := OldRow;
end;

// Charge la Liste2 en fonction de l'axe et la sous-section sélectionnée
procedure TOF_CPSTRUCTURE.TSSPlanOpen;
var
  OldRow : Integer;
begin
  OldRow := FListe2.Row;
  TSSPlan.LoadDetailDB('SSSTRUCR','"'+FAxe+'";"'+FListe1.Cells[2,FListe1.Row]+'"','',nil,False,False);
  TSSPlan.PutGridDetail(FListe2,False,False,'PS_AXE;PS_SOUSSECTION;PS_CODE;PS_LIBELLE;PS_ABREGE;PS_SOCIETE;PS_CORRESP;PS_TOTDEBP;PS_TOTCREP;PS_TOTDEBE;PS_TOTCREE;PS_TOTDEBS;PS_TOTCRES;PS_TOTDEBANO;PS_TOTCREANO;PS_TOTDEBANON1;PS_TOTCREANON1;PS_DOMAINE',True);
  THPanel(GetControl('PSECTION',True)).Caption := MsgBox.Mess[4]+FListe1.Cells[2,FListe1.Row];
  THPanel(GetControl('PPLAN',True)).Caption := MsgBox.Mess[2]+FListe1.Cells[3,FListe1.Row];
  MemoLg := StrToInt('0'+FListe1.Cells[6,FListe1.Row]);

  ActiveBouton;
  if (FListe2.RowCount > OldRow) then FListe2.Row := OldRow;
end;

// Inverse la valeur de la colonne contrôle de la Liste1 ('X' ou '-')
procedure TOF_CPSTRUCTURE.SwapLaValeur;
var
  St : String ;
begin
  if FListe1.Cells[4,FListe1.Row] = '-' then FListe1.Cells[4,FListe1.Row] := 'X'
                                        else FListe1.Cells[4,FListe1.Row] := '-';

  // Empêche de modifier la valeur si le code Plan n'est pas renseigné (en création d'enregistrement)
  if FListe1.ColEditables[2] then Exit;

  TPlan.Detail[FListe1.Row-1].PutValue('SS_CONTROLE',FListe1.Cells[4,FListe1.Row]);
  TPlan.Detail[FListe1.Row-1].UpdateDB;

  St := FListe1.Cells[2,FListe1.Row];
  ExecuteSql('UPDATE STRUCRSE SET SS_CONTROLE="-" WHERE SS_AXE="'+FAxe+'" AND '+'SS_SOUSSECTION<>"'+St+'" ') ;
  TPlanOpen(True);
end;

procedure TOF_CPSTRUCTURE.FListe1DblClick(Sender: TObject);
var
  OldRow : Integer;
begin
  OldRow := FListe1.Row;
  if (FListe1.Col = 4) then SwapLaValeur;
  FListe1.Row := OldRow;
end;

procedure TOF_CPSTRUCTURE.FListe1RowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  TSSPlanOpen;
end;

// Sauvegarde les modifications de la lignes
procedure TOF_CPSTRUCTURE.FListe1RowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var
  OldRow : Integer;
begin
  // Si création d'enregistrement : Valide
  if FListe1.ColEditables[2] then begin
    FListe1.Row := Ou;
    Cancel := True;
    BValider1Click(nil);
    Exit;
  end;

  OldRow := FListe1.Row;
  FListe1.Row := Ou;
  Bouge1(nbPost);
  FListe1.Row := OldRow;
end;

procedure TOF_CPSTRUCTURE.FListe1KeyPress(Sender: TObject; var Key: Char);
begin
  if (FListe1.Col = 4) then begin
    if Key=#32 then SwapLaValeur;
    Key := #0;
    // Permet d'afficher la coche et non un X
    FListe1.Col := 3;
    FListe1.Col := 4;
  end;
end;

procedure TOF_CPSTRUCTURE.FListe1KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if ((ssCtrl in Shift) and (Key=VK_DELETE) and (TPlan.Detail.Count = 0)) then begin Key:=0; Exit; end;
  if ((ssCtrl in Shift) and (Key=VK_DELETE))then begin Bouge1(nbDelete); Key:=0; end;
end;

procedure TOF_CPSTRUCTURE.FListe2KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if ((ssCtrl in Shift) and (Key=VK_DELETE) and (TSSPlan.Detail.Count = 0)) then begin Key:=0; Exit; end;
  if ((ssCtrl in Shift) and (Key=VK_DELETE))then begin Bouge2(nbDelete); Key:=0; end;
end;

procedure TOF_CPSTRUCTURE.FListe2RowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var
  OldRow : Integer;
begin
  // Si création d'enregistrement : Valide
  if FListe2.ColEditables[3] then begin
    FListe2.Row := Ou;
    Cancel := True;
    BValiderClick(nil);
    Exit;
  end;

  OldRow := FListe2.Row;
  FListe2.Row := Ou;
  Bouge2(nbPost);
  FListe2.Row := OldRow;
end;

function TOF_CPSTRUCTURE.Bouge1(Button: TNavigateBtn): boolean;
begin
  Result := False;
  Case Button of
    nbLast,nbPrior,nbNext,
    nbFirst,nbInsert : if Not OnSauve1 then Exit;
    nbPost           : if Not EnregOK1 then Exit;
    nbDelete         : if MsgBox1.Execute(1,'','')<>mrYes then Exit; // Confirmez-vous la suppression de l'enregistrement ?
    nbCancel         : begin FListe1.ColEditables[2] := False; FListe1.ColEditables[4] := False; TPlanOpen; TSSPlanOpen; end;
    nbRefresh        : begin TPlanOpen; TSSPlanOpen; end;
  end;
  if (Button=nbInsert) then if Not NewEnreg1Possible then Exit;

  // Sauvegarde
  if (Button=nbPost) then begin
    // Si création d'enregistrement: Ajoute une ligne à la tob
    if FListe1.ColEditables[2] then Tob.Create('STRUCRSE',TPlan,FListe1.Row-1);

    if TPlanModified then begin
      if not (TPlan.Detail[FListe1.Row-1].InsertOrUpdateDB) then begin MessageAlerte(Msgbox.mess[1]); Exit; end;  // L'enregistrement est inaccessible
      TPlan.Detail.Sort('SS_SOUSSECTION');
      FListe1.ColEditables[2] := False;
      FListe1.ColEditables[4] := False;
      Bouge1(nbRefresh);
      FAvertir:=True;
    end;
  end;
  // Suppression
  if (Button=nbDelete) then begin
    if not (TPlan.Detail[FListe1.Row-1].DeleteDB(False)) then begin MessageAlerte(Msgbox.mess[1]); Exit; end; // L'enregistrement est inaccessible
    TPlan.Detail[FListe1.Row-1].Free;
    FAvertir:=True;
    FListe1.DeleteRow(FListe1.Row);
    TSSPlanOpen;
  end;

  Result := True;
  if (Button=nbInsert) then NewEnreg1;
  ActiveBouton;
end;

function TOF_CPSTRUCTURE.EnregOK1: boolean;
begin
  Result := False;

  if Not ValideLg then begin
    FListe1.SetFocus;
    FListe1.Col := 6;
    Exit;
  end;

  if FListe1.Cells[2,FListe1.Row] = '' then begin
    MsgBox1.Execute(2,'',''); // Vous devez renseigner un code.
    FListe1.SetFocus;
    FListe1.Col := 2;
    Exit;
  end;

  if (FListe1.Cells[3,FListe1.Row] = '') then begin
    MsgBox1.Execute(3,'',''); // Vous devez renseigner un libellé.
    FListe1.SetFocus;
    FListe1.Col := 3;
    Exit;
  end;

  // Si c'est une création d'enregistrement
  if FListe1.ColEditables[2] then begin
    if VerifiSiExiste1 then begin
      MsgBox1.Execute(4,'',''); // Le code que vous avez saisi existe déjà. Vous devez le modifier.
      FListe1.SetFocus;
      FListe1.Col := 2;
      Exit;
    end;
  end;
  Result := True;
end;

function TOF_CPSTRUCTURE.OnSauve1: boolean;
var
  Rep : Integer ;
begin
  Result := False;
  if TPlanModified then begin
    // Voulez-vous enregistrer les modifications ?
    Rep := MsgBox1.Execute(0,'','') ;
    end
  else Rep:=321;

  if FClosing then FLastRep := Rep;

  Case Rep of
    mrYes : if not Bouge1(nbPost)   then Exit ;
    mrNo  : if not Bouge1(nbCancel) then Exit ;
    mrCancel : begin Abort ; Exit ; end;
  end ;
  Result := True;
end;

procedure TOF_CPSTRUCTURE.NewEnreg1;
var
  iRow : Integer;
begin
  iRow := FListe1.Row;
  FListe1.InsertRow(FListe1.Row);
  FListe1.Row := iRow;
  FListe1.ColEditables[2] := True;
  FListe1.ColEditables[4] := True;
  FListe1.Cells[1,FListe1.Row] := FAxe;              // SS_AXE
  FListe1.Cells[4,FListe1.Row] := '-';               // SS_CONTROLE
  FListe1.Cells[5,FListe1.Row] := IntToStr(LgDeb);   // SS_DEBUT
  FListe1.Cells[6,FListe1.Row] := '0';               // SS_LONGUEUR
  FListe1.Cells[7,FListe1.Row] := V_PGI.CodeSociete; // SS_SOCIETE
  FListe1.Col := 2;

  TSSPlanOpen;

  FListe1.SetFocus;
end;

function TOF_CPSTRUCTURE.NewEnreg1Possible: Boolean;
var
  Lg : Byte ;
  i : Byte ;
begin
  Result := True;
  Lg := SommeLongeur;
  if (Lg>=FLgSect) then begin
    MsgBox1.Execute(5,'',''); // Vous avez déjà atteint la longueur maximum des sections de cet axe.
    Result := False;
    Bouge1(nbCancel);
    end
  else begin
    i := ChercheUnTrou;
    if (i=0) then LgDeb := Lg+1
             else LgDeb := i;
  end;

{$IFDEF CCS3}
  If Result And (NbSousPlan=3) then begin
    MsgBox1.Execute(12,'',''); // Vous ne pouvez pas renseigner plus de 3 sous-plans.
    Result := False;
    Bouge1(nbCancel);
  end;
{$ENDIF}
end;

// Regarde si on a modifié quelque chose
function TOF_CPSTRUCTURE.TPlanModified: Boolean;
begin
  Result := False;
  if (TPlan.Detail.Count = 0) then Exit;

  // Met à jour la TOB avec la ligne
  TPlan.Detail[FListe1.Row-1].GetLigneGrid(FListe1,FListe1.Row,';SS_AXE;SS_SOUSSECTION;SS_LIBELLE;SS_CONTROLE;SS_DEBUT;SS_LONGUEUR;SS_SOCIETE');

  // Regard si la TOB a été modifiée
  Result := TPlan.Detail[FListe1.Row-1].IsOneModifie(False);
end;

function TOF_CPSTRUCTURE.TSSPlanModified: Boolean;
begin
  Result := False;
  if (TSSPlan.Detail.Count = 0) then Exit;

  // Met à jour la TOB avec la ligne
  TSSPlan.Detail[FListe2.Row-1].GetLigneGrid(FListe2,FListe2.Row,';PS_AXE;PS_SOUSSECTION;PS_CODE;PS_LIBELLE;PS_ABREGE;PS_SOCIETE;PS_CORRESP;PS_TOTDEBP;PS_TOTCREP;PS_TOTDEBE;PS_TOTCREE;PS_TOTDEBS;PS_TOTCRES;PS_TOTDEBANO;PS_TOTCREANO;PS_TOTDEBANON1;PS_TOTCREANON1;PS_DOMAINE');

  // Regard si la TOB a été modifiée
  Result := TSSPlan.Detail[FListe2.Row-1].IsOneModifie(False);
end;

// Vérifie si la TOB en cours de création n'a pas le même code qu'une déja existante
function TOF_CPSTRUCTURE.VerifiSiExiste1: Boolean;
var
  TF : Tob;
begin
  TF := TPlan.FindFirst(['SS_AXE','SS_SOUSSECTION'],[FAxe,FListe1.Cells[2,FListe1.Row]],False);
  if (TF = nil) then Result := False
                else Result := True;
end;

function TOF_CPSTRUCTURE.ValideLg: Boolean;
var
  i : Integer ;
  iValeur : Integer;
  szValeur : String;
BEGIN
  Result := True;

  if (TPlan.Detail.Count = 0) then Exit;

  szValeur := Trim(FListe1.Cells[6,FListe1.Row]);
  if (szValeur = '') then begin
    Result := False;
    MsgBox1.Execute(13,'',''); // Vous devez renseigner une longueur.
    Exit;
  end;

  iValeur := StrToInt(szValeur);
  if (iValeur = 0) then begin
    Result := False;
    MsgBox1.Execute(11,'',''); // Vous ne pouvez pas renseigner un longueur égale à 0.
    Exit;
  end;

  if (iValeur > MemoLg) then begin
    i := SommeLongeur;
    i := i-MemoLg + iValeur;
    if (i > FLgSect) then begin
      Result := False;
      MsgBox1.Execute(6,'',''); // La longueur que vous avez renseignée est trop grande.
      Exit;
    end;
  end;

  if (iValeur<>MemoLg) then begin
    if Not ChercheEnreg2 then begin
      if MsgBox1.Execute(7,'','')=mrYes then DetruitContenuPlan // La modification de la longueur va entraîner la destruction du contenu des plans de sous-sections. Désirez-vous continuez ?
                                        else begin FListe1.Cells[6,FListe1.Row] := IntToStr(MemoLg); Result := False; end;
    end;
  end;
end;

function TOF_CPSTRUCTURE.SommeLongeur: Integer;
var
  Q : TQuery ;
begin
  Result :=0;
  Q := OpenSql('SELECT SUM(SS_LONGUEUR) FROM STRUCRSE WHERE SS_AXE="'+FAxe+'"',True);
  If Not Q.Eof then Result := Q.Fields[0].AsInteger; Ferme(Q);
end;

function TOF_CPSTRUCTURE.ChercheEnreg2: Boolean;
var
  Q : TQuery ;
begin
  Q := OpenSql('SELECT PS_CODE FROM SSSTRUCR WHERE PS_AXE="'+FAxe+'" AND '+'PS_SOUSSECTION="'+FListe1.Cells[2,FListe1.Row]+'"',True);
  Result := (Q.Eof);
  Ferme(Q);
end;

procedure TOF_CPSTRUCTURE.DetruitContenuPlan;
begin
  ExecuteSQL('DELETE FROM SSSTRUCR WHERE PS_AXE="'+FAxe+'" AND '+'PS_SOUSSECTION="'+FListe1.Cells[2,FListe1.Row]+'"') ;
  TSSPlanOpen;
end;

function TOF_CPSTRUCTURE.Bouge2(Button: TNavigateBtn): boolean;
begin
  Result := False;
  Case Button of
    nbLast,nbPrior,nbNext,
    nbFirst,nbInsert : if Not OnSauve2 then Exit;
    nbPost           : if Not EnregOK2 then Exit;
    nbDelete         : if MsgBox2.Execute(1,'','')<>mrYes then Exit; // Confirmez-vous la suppression de l'enregistrement ?
    nbCancel         : begin FListe2.ColEditables[3] := False; TSSPlanOpen; end;
    nbRefresh        : begin TSSPlanOpen; end;
  end;

  // Sauvegarde
  if (Button=nbPost) then begin
    // Si création d'enregistrement: Ajoute une ligne à la tob
    if FListe2.ColEditables[3] then Tob.Create('SSSTRUCR',TSSPlan,FListe2.Row-1);

    if TSSPlanModified then begin
      if not (TSSPlan.Detail[FListe2.Row-1].InsertOrUpdateDB) then begin MessageAlerte(Msgbox.mess[1]); Exit; end;  // L'enregistrement est inaccessible
      TSSPlan.Detail.Sort('PS_CODE');
      FListe2.ColEditables[3] := False;
      Bouge2(nbRefresh);
      FAvertir:=True;
    end;
  end;
  // Suppression
  if (Button=nbDelete) then begin
    if not (TSSPlan.Detail[FListe2.Row-1].DeleteDB(False)) then begin MessageAlerte(Msgbox.mess[1]); Exit; end; // L'enregistrement est inaccessible
    TSSPlan.Detail[FListe2.Row-1].Free;
    FAvertir:=True;
    FListe2.DeleteRow(FListe2.Row);
    TSSPlanOpen;
  end;

  Result := True;
  if (Button=nbInsert) then NewEnreg2;
  ActiveBouton;
end;

function TOF_CPSTRUCTURE.OnSauve2: boolean;
var
  Rep : Integer ;
begin
  Result := False;

  if FClosing then
    Rep := FLastRep
  else begin
    if TSSPlanModified then begin
      // Voulez-vous enregistrer les modifications ?
      Rep:=MsgBox2.execute(0,'','');
      end
    else rep:=321 ;
  end;

  Case rep of
    mrYes : if not Bouge2(nbPost)   then Exit;
    mrNo  : if not Bouge2(nbCancel) then Exit;
    mrCancel : begin Abort; Exit; end;
  end;
  Result := True;
end;

procedure TOF_CPSTRUCTURE.NewEnreg2;
var
  iRow : Integer;
begin
  if (OnlyOne) And (TSSPlan.Detail.Count <> 0) then begin
    Bouge2(nbCancel);
    Exit ;
  end;

  iRow := FListe2.Row;
  FListe2.InsertRow(FListe2.Row);
  FListe2.Row := iRow;

  FListe2.ColEditables[3] := True;
  FListe2.ColLengths[3] := StrToInt(FListe1.Cells[6,FListe1.Row])  ;

  FListe2.Cells[1,FListe2.Row] := FAxe;                         // PS_AXE
  FListe2.Cells[2,FListe2.Row] := FListe1.Cells[2,FListe1.Row]; // PS_SOUSSECTION
  FListe2.Cells[5,FListe2.Row] := '';                           // PS_ABREGE
  FListe2.Cells[6,FListe2.Row] := V_PGI.CodeSociete;            // PS_SOCIETE
  FListe2.Cells[7,FListe2.Row] := '';                           // PS_CORRESP
  FListe2.Cells[8,FListe2.Row] := '0';                          // PS_TOTDEBP
  FListe2.Cells[9,FListe2.Row] := '0';                          // PS_TOTCREP
  FListe2.Cells[10,FListe2.Row] := '0';                         // PS_TOTDEBE
  FListe2.Cells[11,FListe2.Row] := '0';                         // PS_TOTCREE
  FListe2.Cells[12,FListe2.Row] := '0';                         // PS_TOTDEBS
  FListe2.Cells[13,FListe2.Row] := '0';                         // PS_TOTCRES
  FListe2.Cells[14,FListe2.Row] := '0';                         // PS_TOTDEBANO
  FListe2.Cells[15,FListe2.Row] := '0';                         // PS_TOTCREANO
  FListe2.Cells[16,FListe2.Row] := '0';                         // PS_TOTDEBANON1
  FListe2.Cells[17,FListe2.Row] := '0';                         // PS_TOTCREANON1
  FListe2.Cells[18,FListe2.Row] := '';                          // PS_DOMAINE

  FListe2.Col := 3;

  FListe2.SetFocus;
end;

function TOF_CPSTRUCTURE.EnregOK2: boolean;
begin
  Result := False;

  if FListe2.Cells[3,FListe2.Row] = '' then begin
    MsgBox2.Execute(2,'',''); // Vous devez renseigner un code.
    FListe2.Col := 3;
    FListe2.SetFocus;
    Exit;
  end;

  if FListe2.Cells[4,FListe2.Row] = '' then begin
    MsgBox2.Execute(3,'','') ; // Vous devez renseigner un libellé.
    FListe2.Col := 4;
    FListe2.SetFocus;
    Exit;
  end;

  if Not BourreLeCompte then Exit;

  // Si c'est une création d'enregistrement
  if FListe2.ColEditables[3] then begin
    if VerifiSiExiste2 then begin
      MsgBox2.Execute(4,'',''); // Le code que vous avez saisi existe déjà. Vous devez le modifier.
      FListe2.Col := 3;
      FListe2.SetFocus;
      Exit;
    end;
  end;
  Result := True;
end;

function TOF_CPSTRUCTURE.VerifiSiExiste2: Boolean;
var
  TF : Tob;
begin
  TF := TSSPlan.FindFirst(['PS_SOUSSECTION'],[FListe2.Cells[3,FListe2.Row]],False);
  if (TF = nil) then Result := False
                else Result := True;
end;

function TOF_CPSTRUCTURE.ChercheUnTrou: Byte;
var
  Trou : Byte;
  Q : TQuery;
  Premiertour : Boolean;
begin
  if (TPlan.Detail.Count = 0) then begin
    Result := 0;
    Exit;
  end;

  Trou:=0;
  Q := OpenSql('SELECT SS_DEBUT,SS_LONGUEUR FROM STRUCRSE WHERE SS_AXE="'+FAxe+'" '+'ORDER BY SS_DEBUT',True);
  PremierTour := True;

  While Not Q.Eof do begin
    if PremierTour then begin
      if Q.Fields[0].AsInteger<>1 then begin
        Trou:=1;
        Break;
        end
      else
        PremierTour:=False ;
    end;
    Trou := Q.Fields[0].AsInteger + Q.Fields[1].AsInteger;
    Q.Next;
    if (Trou<>Q.Fields[0].AsInteger) and (Not Q.Eof) then Break;
  end;
  if Q.Eof then Trou := 0;
  Ferme(Q);
  Result := Trou;
end;

function TOF_CPSTRUCTURE.TrouDansEnchainement(var NumMess: Byte): Boolean;
var
  Q : TQuery ;
  Ok : Boolean ;
  DL : Byte ;
  szValeur : String;
  iValeur : Integer;
begin
  if (TPlan.Detail.Count = 0) then begin
    Result := False;
    Exit;
  end;

  szValeur := Trim(FListe1.Cells[6,FListe1.Row]);
  if (szValeur = '') then begin
    NumMess := 13;  // Vous devez renseigner une longueur.
    Result := True;
    Exit;
  end;

  iValeur := StrToInt(szValeur);
  if (iValeur = 0) then begin
    NumMess := 11; // Vous ne pouvez pas renseigner une longueur égale à 0.
    Result := True;
    Exit;
  end;

  if (iValeur < 0) then begin
    NumMess := 9; // Une des longueurs que vous avez renseignées n'est pas valide.
    Result := True;
    Exit;
  end;

  Q := OpenSql('SELECT SS_DEBUT,SS_LONGUEUR FROM STRUCRSE WHERE SS_AXE="'+FAxe+'" '+ 'ORDER BY SS_DEBUT',True);
  Ok := False;

  While Not Q.Eof And Not Ok do begin
    DL := Q.Fields[0].AsInteger + Q.Fields[1].AsInteger;
    Q.Next;
    if (DL<>Q.Fields[0].AsInteger) and (Not Q.Eof) then begin
      Ok := True;
      NumMess := 8; // Votre axe est mal structuré : il n'y a pas de continuité dans les enchaînements début longueur.
//      FListe1.Columns.Items[3].ReadOnly:=False ;
    end;
  end;
  Ferme(Q);
  Result := Ok;
end;

{$IFDEF CCS3}
function TOF_CPSTRUCTURE.NbSousPlan: Integer;
var
  Q : TQuery ;
begin
  Result := 0;
  Q := OpenSql('SELECT COUNT(*) FROM STRUCRSE WHERE SS_AXE="'+FAxe+'"',True);
  if Not Q.Eof then
    Result := Q.Fields[0].AsInteger;
  Ferme(Q);
end;
{$ENDIF}

function TOF_CPSTRUCTURE.BourreLeCompte: Boolean;
var
  i : Integer ;
  iLongueur : Integer;
  szCode : String;
begin
  Result := True;

  if MauvaiseLg then begin
    if not OnlyOne then begin
      szCode := FListe2.Cells[3,FListe2.Row];
      { FQ 21322 BVE 03.09.07
      iLongueur := StrToInt(FListe2.Cells[4,FListe2.Row]); }
      iLongueur := StrToInt(FListe1.Cells[6,FListe1.Row]);
      for i := Length(szCode)+1 to iLongueur do
        szCode := szCode + VH^.Cpta[AxeToFb(Faxe)].Cb;
      FListe2.Cells[3,FListe2.Row] := szCode;
      end
    else begin
      MsgBox2.execute(6,'',''); // Le code choisi ne respecte pas la longueur de cette sous-section.
      FListe2.Col := 1;
      FListe2.SetFocus;
      Result := False;
      Exit;
    end;
  end;
end;

function TOF_CPSTRUCTURE.MauvaiseLg: Boolean;
begin
  Result := (Length(FListe2.Cells[3,FListe2.Row]) < StrToInt(FListe1.Cells[6,FListe1.Row]));
end;

function TOF_CPSTRUCTURE.OnlyOne: Boolean;
begin
  if (FListe1.Cells[4,FListe1.Row] = 'X') then Result := True
                                          else Result := False;
end;

procedure TOF_CPSTRUCTURE.GenereCodeContenuPlan;
var
  Q : TQuery ;
begin
  Q := OpenSql('SELECT S_SECTION FROM SECTION WHERE S_AXE="'+FAxe+'"',True);
  InitMove(RecordsCount(Q),MsgBox2.Mess[7]); // Génération des codes
  (* BVE 02.05.07
{$IFNDEF EAGLCLIENT}
  QStruc := TQuery.Create(Application) ;
  QStruc.DataBaseName:='SOC' ;
{$ENDIF}
  END BVE 02.05.07 *)
  BeginTrans ;
  try
    While Not Q.Eof do begin
      CodeStructure(Faxe,Q.Fields[0].AsString);
      Q.Next;
      MoveCur(False);
    end;
    CommitTrans;
  except
    Rollback;
  end;
  Ferme(Q);
  FiniMove;
end;

procedure TOF_CPSTRUCTURE.ActiveBouton;
begin
 SetControlEnabled('BDEFAIRE', not FListe2.ColEditables[3]);
 SetControlEnabled('BINSERT', (not FListe1.ColEditables[2]) and (not FListe2.ColEditables[3]));
 SetControlEnabled('BDELETE', (not FListe1.ColEditables[2]) and (not FListe2.ColEditables[3]) and (TPlan.Detail.Count > 0));

 SetControlEnabled('BDEFAIRE1', FListe2.ColEditables[3] or ( (not FListe1.ColEditables[2]) and (not FListe2.ColEditables[3])) );
 SetControlEnabled('BINSERT1', (not FListe1.ColEditables[2]) and (not FListe2.ColEditables[3]) and (TPlan.Detail.Count > 0));
 SetControlEnabled('BDELETE1', (not FListe1.ColEditables[2]) and (not FListe2.ColEditables[3]) and (TSSPlan.Detail.Count > 0));

 SetControlEnabled('BFERME', not FListe1.ColEditables[2]);

 SetControlEnabled('BMULTISECT', (not FListe1.ColEditables[2]) and (not FListe2.ColEditables[3]));
 SetControlEnabled('BSIMPSECT',  (not FListe1.ColEditables[2]) and (not FListe2.ColEditables[3]));

// if (TSSPlan.Detail.Count = 0) then SetControlEnabled('BDELETE1', False);
 if (TPlan.Detail.Count = 0) or (TPlan.Detail.Count = 1) or (FListe1.Cells[2,FListe1.Row] = '') or (FListe1.ColEditables[2]) or (FListe2.ColEditables[3]) then begin
   SetControlEnabled('BAUTOMATE', False);
   SetControlEnabled('BMULTISECT', False);
   SetControlEnabled('BSIMPSECT', False);
   end
 else begin
   SetControlEnabled('BAUTOMATE', True);
   SetControlEnabled('BMULTISECT', True);
   SetControlEnabled('BSIMPSECT', True);
 end;
end;

Initialization
  registerclasses ( [ TOF_CPSTRUCTURE ] ) ;
end.
