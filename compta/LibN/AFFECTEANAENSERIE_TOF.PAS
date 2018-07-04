{***********UNITE*************************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/11/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFFECTEANAENSERIE ()
Mots clefs ... : TOF;AFFECTEANAENSERIE
*****************************************************************}
Unit AffecteAnaEnSerie_Tof ;

Interface

Uses StdCtrls, 
     Controls,
     Classes,
     Graphics, 
{$IFDEF EAGLCLIENT}
     eMul,
     MainEagl, // AGLLanceFiche
{$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     FE_main,  // AGLLanceFiche
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     LookUp,
     AGLInit,   // TheTOB
     uTobDebug, // TobDebug
     uTob,      // Tob
     Ent1,      // MaxAxe
     hTB97,     // TToolBarButton97
     Windows,   // VK
     UTOF ;

Type
  TOF_AFFECTEANAENSERIE = Class (TOF)

    RG         : THRadioGroup;
    Axe        : THValComboBox;
    Section    : THEdit;
    VentilType : THRadioGroup;

    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;

    procedure OnClickRG             ( Sender : TObject );
    procedure OnChangeAxe           ( Sender : TObject );
    procedure OnElipsisClickSection ( Sender : TObject );
    procedure OnKeyDownSection      ( Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure OnExitSection         ( Sender : TObject );

  private
    FStAxe  : string;   // Ax
    FNumAxe : integer;  // 1,2,3,4,5

    procedure GestionComposant;
  end ;

function CPLanceFiche_AffecteAnalytiqueEnSerie : string;

Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  {$ENDIF MODENT1}
  uLibWindows,
  uLibAnalytique, // AlloueAxe
  SaisUtil;       // EcrVersAna

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/11/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function CPLanceFiche_AffecteAnalytiqueEnSerie : string;
begin
  Result := AGLLanceFiche('CP', 'CPAFFECTEANA', '', '', '');
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/11/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_AFFECTEANAENSERIE.OnArgument (S : String ) ;
var lStGeneral : string;
    lStPlus : string;
    lQuery : TQuery;
    i : integer;
begin
  Inherited ;
  RG         := THRadioGroup(GetControl('RG', True));
  Axe        := THValComboBox(GetControl('AXE', True));
  Section    := THEdit(GetControl('SECTION', True));
  VentilType := THRadioGRoup(GetControl('VENTILTYPE', True));

  Axe.OnChange := OnChangeAxe;
  RG.OnClick   := OnClickRG;
  Section.OnElipsisClick := OnElipsisClickSection;
  Section.OnKeyDown      := OnKeyDownSection;
  Section.OnExit         := OnExitSection;

  //
  if (LaTob <> nil) and (LaTob.Detail.Count <> 0) then
  begin
    lStGeneral := laTob.Detail[0].GetString('E_GENERAL');

    try
      lQuery := OpenSql('SELECT G_VENTILABLE1, G_VENTILABLE2, G_VENTILABLE3, ' +
                        'G_VENTILABLE4, G_VENTILABLE5 FROM GENERAUX WHERE ' +
                        'G_GENERAL = "' + lStGeneral + '"', True);

      for i := 1 to MaxAxe do
      begin
        if lQuery.FindField('G_VENTILABLE'+ IntToStr(i)).AsString = 'X' then
        begin
          if lStPlus <> '' then
            lStPlus := lStPlus + ' OR ';

          lStPlus := lStPlus + ' X_AXE = "A' + IntToStr(i) + '"';
        end;
      end;

    finally
      Ferme(lQuery);
      Axe.Plus := ' AND (' + lStPlus + ')';
    end;
  end;

  Axe.ItemIndex := 0;
  RG.ItemIndex  := 0;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/11/2005
Modifié le ... : 12/07/2006
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_AFFECTEANAENSERIE.OnLoad ;
begin
  inherited ;
  // GCO - 12/07/2006 - FQ 17495
  if Axe.CanFocus then
    Axe.SetFocus;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/11/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_AFFECTEANAENSERIE.OnDisplay () ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/11/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_AFFECTEANAENSERIE.OnUpdate ;
var 
    lStSection : string;
    lStSql : string;
    lNumVentil : integer;
    lPourcentage : Double;
    i,j : integer;
    lTobEcr, lTobAna, lTobAxe : Tob;
    lQuery : TQuery;
begin
  Inherited ;

  FStAxe  := Axe.Value;
  FNumAxe := StrToInt(Copy(FStAxe, 2, 1));

  if Section.Enabled then
  begin
    if not Presence('SECTION', 'S_SECTION', Section.Text) then
    begin
      PgiInfo('La section n''existe pas.', Ecran.Caption);
      Exit;
    end;
  end;

  for i := 0 to LaTob.Detail.Count-1 do
  begin
    lTobEcr := LaTob.Detail[i];

    // Création des Axes
    AlloueAxe(lTobEcr);

    lTobAxe := lTobEcr.Detail[FNumAxe-1];

    lStSql := 'DELETE FROM ANALYTIQ WHERE ' +
              'Y_JOURNAL = "' + lTobEcr.GetString('E_JOURNAL') + '" AND ' +
              'Y_EXERCICE = "' + lTobEcr.GetString('E_EXERCICE') + '" AND ' +
              'Y_DATECOMPTABLE = "' + UsDateTime(lTobEcr.GetDateTime('E_DATECOMPTABLE')) + '" AND ' +
              'Y_NUMEROPIECE = ' + IntToStr(lTobEcr.GetInteger('E_NUMEROPIECE')) + ' AND ' +
              'Y_NUMLIGNE = ' + IntToStr(lTobEcr.GetInteger('E_NUMLIGNE')) + ' AND ' +
              'Y_GENERAL = "' + lTobEcr.GetString('E_GENERAL') + '" AND ' +
              'Y_AXE = "' + FStAxe + '"';

    ExecuteSQL(lStSql);

    if Section.Enabled then
    begin
      lTobAna := TOB.Create('ANALYTIQ', lTobAxe, -1) ;
      EcrVersAna(lTobEcr, lTobAna) ;
      VentilLigneTOB( lTobAna, Section.Text, 1, V_Pgi.OkDecV, 100, lTOBEcr.GetDouble('E_DEBIT')<> 0) ;
      lTobAna.InsertDb( nil, False);
      FreeAndNil(lTobAna);
    end
    else
    begin
      lQuery := OpenSql('SELECT * FROM VENTIL WHERE ' +
                        'V_NATURE = "TY' + IntToStr(FNumAxe) + '" AND ' +
                        'V_COMPTE = "' + GetControlText('VENTILTYPE') + '" ' +
                        'ORDER BY V_NATURE, V_COMPTE, V_NUMEROVENTIL',True) ;

      while not lQuery.Eof do
      begin
        lTOBAna := TOB.Create('ANALYTIQ', lTobAxe, -1);
        EcrVersAna( lTobEcr, lTobAna) ;
        lStSection   := lQuery.FindField('V_SECTION').AsString;
        lNumVentil   := lQuery.FindField('V_NUMEROVENTIL').AsInteger;
        lPourcentage := lQuery.FindField('V_TAUXMONTANT').AsFloat ;

        VentilLigneTOB( lTOBAna, lStSection, lNumVentil, V_Pgi.OkDecV,
                        lPourcentage, lTobEcr.GetDouble('E_DEBIT') <> 0,
                        lQuery.FindField('V_TAUXQTE1').AsFloat,
                        lQuery.FindField('V_TAUXQTE2').AsFloat);
        lQuery.Next ;
      end;

      Ferme(lQuery);

      ArrondirAnaTOB(lTobEcr, V_Pgi.OkDecV);


      for j := 0 to lTobAxe.Detail.Count -1 do
      begin
        lTobAxe.Detail[j].InsertDb( nil, False );
      end;
    end;
    lTobEcr.ClearDetail;
  end;

  Ecran.Close;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/11/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_AFFECTEANAENSERIE.OnClose ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/11/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_AFFECTEANAENSERIE.OnClickRG(Sender: TObject);
begin
  GestionComposant;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/11/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_AFFECTEANAENSERIE.OnChangeAxe(Sender: TObject);
begin
  GestionComposant;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/11/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_AFFECTEANAENSERIE.GestionComposant;
begin
  Section.Enabled    := (Rg.ItemIndex = 0);
  VentilType.Enabled := (Rg.ItemIndex = 1);

  Section.Color    := IIF( Section.Enabled, clWindow, ClBtnFace);
  VentilType.Color := IIF( VentilType.Enabled, clWindow, ClBtnFace);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/11/2005
Modifié le ... :   /  /    
Description .. :

Mots clefs ... : 
*****************************************************************}
procedure TOF_AFFECTEANAENSERIE.OnElipsisClickSection(Sender: TObject);
begin
  LookUpList(THEdit(Sender), TraduireMemoire('Sections'), 'SECTION', 'S_SECTION', 'S_LIBELLE', 'S_AXE = "' + Axe.Value + '"', 'S_SECTION', True, 3);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 31/05/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_AFFECTEANAENSERIE.OnKeyDownSection(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
    VK_F5: if (Shift = []) then
           begin
             LookUpList(THEdit(Sender), TraduireMemoire('Sections'), 'SECTION', 'S_SECTION', 'S_LIBELLE', 'S_AXE = "' + Axe.Value + '"', 'S_SECTION', True, 3);
           end
           else
           if (Shift = [ssShift]) then
           begin
             LookUpList(THEdit(Sender), TraduireMemoire('Sections'), 'SECTION', 'S_SECTION', 'S_LIBELLE', 'S_AXE = "' + Axe.Value + '"', 'S_LIBELLE', True, 3);
           end;
  else
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/07/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_AFFECTEANAENSERIE.OnExitSection(Sender: TObject);
begin
  // GCO - 12/07/2006 - FQ 17495
  if (Section.Text <> '') and
     Presence('SECTION"', 'S_SECTION', BourreEtLess(Section.Text, fbSect)) then
  begin
    Section.Text := BourreEtLess(Section.Text, fbSect);
  end
  else
  begin
    OnElipsisClickSection(Sender);
  end;
end;

////////////////////////////////////////////////////////////////////////////////

Initialization
  registerclasses ( [ TOF_AFFECTEANAENSERIE ] ) ; 
end.
