{***********UNITE*************************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 16/03/2001
Modifié le ... : 06/05/2002
Description .. : Gestion de l'assistant de création d'un dossier de paie
Mots clefs ... : PAIE;ASSIST
*****************************************************************}
{
PT1    : 06/05/2002 VG V582 Ajout de la gestion de l'onglet MSA
PT2    : 13/05/2002 VG V582 Rechargement de la tablette etablissement
PT3    : 22/04/2003 VG V_42 Si aucun exercice social n'existe, rentrer en mode
                           création - FQ N°10636
PT4    : 19/05/2003 VG V_42 Gestion de l'aide
PT5    : 26/09/2003 VG V_42 Vérifications en fin de création de dossier
                            FQ N°10643
PT6    : 24/05/2004 PH V_50 Mise en place accès compta CWAS
PT7    : 28/02/2006 EPI V_65 Modification appel exercices sociaux  FQ 12380 et
                            12738
PT8    : 08/06/2006 SB-VG V_65 Adaptation CWAS
PT9    : 14/03/2007 VG V_72 BQ_GENERAL n'est pas forcément unique
}
unit AssistSOPaie;

interface

uses
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  Classes,
  Controls,
  Forms,
  assist,
  hmsgbox,
  StdCtrls,
  HTB97,
  ComCtrls,
  Hctrls,
  ParamSoc,
  Ent1,
  HEnt1,
  UiUtil,
{$IFNDEF EAGLCLIENT}
  FE_Main,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
  MaineAgl,
{$ENDIF}
  FichComm,
{$IFDEF COMPTAPAIE}
  UTOFMULPARAMGEN,
{$ENDIF}
  PgOutils2,
  UtilSoc,
  UTob,
  ImgList,
  HSysMenu,
  ExtCtrls,
  HPanel,
  sysutils,
  PgOutilsTreso, HImgList;

procedure LanceAssistantPaie;

type
TFAssistSOPaie = class(TFAssist)
    PParamSoc: TTabSheet;
    PEtabliss: TTabSheet;
    PExercice: TTabSheet;
    HBienvenue: THLabel;
    GSeparate: TGroupBox;
    bParamCaracter: TToolbarButton97;
    bExercice: TToolbarButton97;
    bEtabCar: TToolbarButton97;
    bParamCom: TToolbarButton97;
    ImageList: THImageList;
    HLabel4: THLabel;
    GroupBox3: TGroupBox;
    HLabel1: THLabel;
    GroupBox2: TGroupBox;
    HLabel2: THLabel;
    GroupBox4: TGroupBox;
    bSalaries: TToolbarButton97;
    bGeneComp: TToolbarButton97;
    ComboEtab: THValComboBox;
    bEtabliss: TToolbarButton97;
    LEtabliss: TLabel;
    bEtabPro: TToolbarButton97;
    bEtabCon: TToolbarButton97;
    bEtabReg: TToolbarButton97;
    bEtabDad: TToolbarButton97;
    bBanque: TToolbarButton97;
    bConvColl: TToolbarButton97;
    bComGen: TToolbarButton97;
    bEtabInter: TToolbarButton97;
    HelpBtn: TToolbarButton97;

    procedure OnBoutonActivationClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    procedure HelpBtnClick(Sender: TObject);
    procedure bAnnulerClick(Sender: TObject);

  private
    { Déclarations privées }
    CurToolbarButton: TToolbarButton97;
    CurToolBarButtonName, Etab: string;
  public
    { Déclarations publiques }
  end;

var
FAssistSOPaie: TFAssistSOPaie;
PP: THpanel;

implementation

{$R *.DFM}

procedure LanceAssistantPaie();
begin
//PT8
FAssistSOPaie:= TFAssistSOPaie.Create(Application);
PP:=FindInsidePanel ;
if PP=Nil then
   begin
   try
      FAssistSOPaie.ShowModal;
   finally
      FAssistSOPaie.Free;
      end;
   end
else
   begin
   InitInside(FAssistSOPaie,PP) ;
   FAssistSOPaie.Show;
   end;
//FIN PT8
end;

procedure TFAssistSOPaie.OnBoutonActivationClick(Sender: TObject);
var
NomEtab, StEtab: string;
QRechEtab: TQuery;
begin
inherited;
if Sender.ClassName <> 'TToolbarButton97' then
   exit;
CurToolBarButtonName:= TToolbarButton97(Sender).Name;
CurToolbarButton:= TToolbarButton97(Sender);

if P.ActivePage = PParamSoc then
   begin
   if CurToolBarButtonName = 'bParamCaracter' then
      begin
//      AGLLanceFiche('PAY','PARAMETRECAR','','','');
      AglLanceFiche('PAY', 'PARAMETRESOC', '', '', '4');
      TToolbarButton97(Sender).ImageIndex:= 0;
      end
   else
   if CurToolBarButtonName = 'bSalaries' then
      begin
//      AGLLanceFiche('PAY','PARAMETRESAL','','','');
      AglLanceFiche ('PAY', 'PARAMETRESOC', '', '', '3');
      TToolbarButton97(Sender).ImageIndex:= 0;
      end
   else
   if CurToolBarButtonName = 'bGeneComp' then
      begin
//      AGLLanceFiche('PAY','PARAMETREGENECOM','','','');
      AglLanceFiche('PAY', 'PARAMETRESOC', '', '', '5');
      TToolbarButton97(Sender).ImageIndex:= 0;
      end
   else
   if CurToolBarButtonName = 'bEtabliss' then
      begin
      AGLLanceFiche('PAY', 'ETABLISSEMENT', '', '', '');
      TToolbarButton97(Sender).ImageIndex:= 0;
      ChargementTablette('ET', ''); //recharge les tablettes
      ComboEtab.ReLoad;
      end
   else
   if CurToolBarButtonName = 'bConvColl' then
      begin
      AglLanceFiche('PAY', 'CONVENTION', '', '', '');
      TToolbarButton97(Sender).ImageIndex:= 0;
      end
   else
   if CurToolBarButtonName = 'bBanque' then
      begin
      FicheBanque_agl('', taModif, 0);
      TToolbarButton97(Sender).ImageIndex:= 0;
      end
   else
   if CurToolBarButtonName = 'bComGen' then
      begin
{$IFDEF COMPTAPAIE}
      CCLanceFiche_MULGeneraux('C;421');
{$ENDIF}
      TToolbarButton97(Sender).ImageIndex:= 0;
      end;
   end
else
if P.ActivePage = PEtabliss then
   begin
   Etab:= ComboEtab.Value;
   StEtab:= 'SELECT ET_LIBELLE'+
            ' FROM ETABLISS WHERE'+
            ' ET_ETABLISSEMENT = "'+Etab+'"';
   QRechEtab:= OpenSql(StEtab, True);
   NomEtab:= QRechEtab.FindField('ET_LIBELLE').AsString;
   Ferme(QRechEtab);
   if CurToolBarButtonName = 'bEtabCar' then
      begin
      AGLLanceFiche('PAY', 'ETABSOCIAL', '', Etab, Etab+';'+NomEtab+';1');
      TToolbarButton97(Sender).ImageIndex:= 0;
      end
   else
   if CurToolBarButtonName = 'bEtabPro' then
      begin
      AGLLanceFiche('PAY', 'ETABSOCIAL', '', Etab, Etab+';'+NomEtab+';2');
      TToolbarButton97(Sender).ImageIndex:= 0;
      end
   else
   if CurToolBarButtonName = 'bEtabCon' then
      begin
      AGLLanceFiche('PAY', 'ETABSOCIAL', '', Etab, Etab+';'+NomEtab+';3');
      TToolbarButton97(Sender).ImageIndex:= 0;
      end
   else
   if CurToolBarButtonName = 'bEtabReg' then
      begin
      AGLLanceFiche('PAY', 'ETABSOCIAL', '', Etab, Etab+';'+NomEtab+';4');
      TToolbarButton97(Sender).ImageIndex:= 0;
      end
   else
   if CurToolBarButtonName = 'bEtabDad' then
      begin
      AGLLanceFiche('PAY', 'ETABSOCIAL', '', Etab, Etab+';'+NomEtab+';5');
      TToolbarButton97(Sender).ImageIndex:= 0;
      end
   else
   if CurToolBarButtonName = 'bEtabInter' then
      begin
      AGLLanceFiche('PAY', 'ETABSOCIAL', '', Etab, Etab+';'+NomEtab+';6');
      TToolbarButton97(Sender).ImageIndex:= 0;
      end;
   end
else
if P.ActivePage = PExercice then
   begin
   if CurToolBarButtonName = 'bExercice' then
      begin
      if not ExisteSQL('SELECT PEX_EXERCICE FROM EXERSOCIAL') then
// PT7
//         AglLanceFiche('PAY', 'EXERCICESOCIAL', '', '', 'ACTION=CREATION')
         AglLanceFiche('PAY', 'EXERSOCIAL', '', '', 'ACTION=CREATION')
      else
// PT7
//         AglLanceFiche('PAY', 'EXERCICESOCIAL', '', '', '');
         AglLanceFiche('PAY', 'MUL_EXERSOCIAL', 'GRILLE=', '', '');
      TToolbarButton97(Sender).ImageIndex:= 0;
      end
   else
   if CurToolBarButtonName = 'bParamCom' then
      begin
      ParamSociete (False, '', 'SCO_COMPTABLES', '', ChargeSocieteHalley,
                    ChargePageSoc, SauvePageSoc, InterfaceSoc, 1105000);
      TToolbarButton97(Sender).ImageIndex:= 0;
      end;
   end;
end;

procedure TFAssistSOPaie.FormShow(Sender: TObject);
begin
inherited;
bAnnuler.Visible:= True;
bFin.Visible:= True;
bFin.Enabled:= False;
end;

procedure TFAssistSOPaie.bSuivantClick(Sender: TObject);
begin
inherited;
if (bSuivant.Enabled) then
   bFin.Enabled:= False
else
   bFin.Enabled:= True;

if P.ActivePage = PEtabliss then
   begin
{$IFNDEF CCS3}
   if ((GetParamSoc('SO_PGMSA') = False) and
      (GetParamSoc('SO_PGINTERMITTENTS') = False)) then
      bEtabInter.Visible:= False
   else
      bEtabInter.Visible:= True;
{$ELSE}
   bEtabInter.Visible:= False;
{$ENDIF}
   end;
end;

procedure TFAssistSOPaie.bFinClick(Sender: TObject);
var
CodeEtab, St, StPlus: string;
QRech: TQuery;
TEtab, TEtabD: TOB;
Erreur, Nb: integer;
begin
inherited;
Erreur:= mrYes;
StPlus:= PGBanqueCP (False); //PT9

St:= 'SELECT ETB_ETABLISSEMENT, ETB_HORAIREETABL, ETB_RIBSALAIRE,'+
     ' BQ_GENERAL'+
     ' FROM ETABCOMPL'+
     ' LEFT JOIN BANQUECP ON'+
     ' ETB_RIBSALAIRE=BQ_GENERAL'+StPlus;
QRech:= OpenSQL(St, TRUE);
TEtab:= TOB.Create('Les etablissements', nil, -1);
TEtab.LoadDetailDB('ETAB', '', '', QRech, False);
Ferme(QRech);

if (TEtab.Detail.Count <> 0) then
   begin
   for Nb:= 0 to TEtab.Detail.Count-1 do
       begin
       TEtabD:= TEtab.Detail[Nb];
       CodeEtab:= TEtabD.GetValue('ETB_ETABLISSEMENT');
       if (TEtabD.GetValue('ETB_HORAIREETABL') = 0) then
          Erreur:= PGIAsk ('L''horaire de l''établissement "'+CodeEtab+'"#13#10'+
                           'n''est pas renseigné.#13#10'+
                           'Voulez-vous quitter l''assistant ?',
                           FAssistSOPaie.Caption);

       if ((TEtabD.GetValue('ETB_RIBSALAIRE') = '') and (Erreur = mrYes)) then
          Erreur:= PGIAsk ('Le RIB de l''établissement "'+CodeEtab+'"#13#10'+
                           'n''est pas renseigné.#13#10'+
                           'Voulez-vous quitter l''assistant ?',
                           FAssistSOPaie.Caption);

       if ((TEtabD.GetValue('BQ_GENERAL') = NULL) and (Erreur = mrYes)) then
          Erreur:= PGIAsk ('La banque de l''établissement "'+CodeEtab+'"#13#10'+
                           'n''est pas créée.#13#10'+
                           'Voulez-vous quitter l''assistant ?',
                           FAssistSOPaie.Caption);
       end;
   end
else
   Erreur:= PGIAsk ('Il n''existe aucun établissement dans le dossier#13#10'+
                    'Voulez-vous quitter l''assistant ?',
                    FAssistSOPaie.Caption);

FreeAndNil (TEtab);
if (Erreur = mrYes) then
//PT8
   begin
   FAssistSOPaie.Close;
   if IsInside(Self) then
      THPanel(PP).CloseInside;
   end;
//FIN PT8
end;

procedure TFAssistSOPaie.bPrecedentClick(Sender: TObject);
begin
inherited;
bSuivant.Enabled := True;
bFin.Enabled := False;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 19/05/2003
Modifié le ... :   /  /
Description .. : Gestion de l'aide
Mots clefs ... : PAIE;PGASSIST
*****************************************************************}
procedure TFAssistSOPaie.HelpBtnClick(Sender: TObject);
begin
CallHelpTopic(Self);
end;

//PT8
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT / Souad BELMAHJOUB
Créé le ...... : 08/06/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE;PGASSIST
*****************************************************************}
procedure TFAssistSOPaie.bAnnulerClick(Sender: TObject);
begin
inherited;
FAssistSOPaie.Close;
if IsInside(Self) then
   THPanel(PP).CloseInside;
end;
//FIN PT8

end.

