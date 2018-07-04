{***********UNITE*************************************************
Auteur  ...... : Joël TRIFILIEFF
Créé le ...... : 31/01/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : GCLANCEUTIL ()
Description .. : Lancement d'utilitaires
Mots clefs ... : TOF;GCLANCEUTIL
*****************************************************************}
Unit GCLANCEUTIL_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     MainEAGL,
{$ELSE}
     FE_Main,
     db,
     dbtables,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     UTOF,
     HTB97,
     Menus;

function GCLanceFiche_LanceUtil(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Type
  TOF_GCLANCEUTIL = Class (TOF)
    private
        tbRazActivite : TToolbarButton97;
        tbRecCptPieces : TToolbarButton97;
        tbAffecteDepot : TToolbarButton97;
        Menu : TPopupMenu;
        mnLance : TMenuItem;
        mnInfo : TMenuItem;
        GBInfos : TGroupBox;
        LstInfos : TListBox;
        // Choix utilitaire
        procedure RazClick(Sender : TObject);
        procedure CalcCptClick(Sender : TObject);
        procedure AffDepClick(Sender : TObject);
        procedure LanceClick(Sender: TObject);
        procedure InfoClick(Sender: TObject);
        procedure CacheInfo;
        procedure Lancement(Sender: TObject);

    public
        // Procédures de la TOF
        procedure OnArgument (S : String ); override;
        procedure OnClose;                  override;
  end ;

Implementation

uses AssistRazActivite,
     AssistAffecteDepot,
     AssistMajCompteurSouche,
     UTOB;

var NumChoix : integer;

function GCLanceFiche_LanceUtil(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
    Result := '';
    if Nat = '' then exit;
    if Cod = '' then exit;
    Result := AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

{===============================================================================
 ====================== Procédures de la TOF ===================================
 ===============================================================================}
procedure TOF_GCLANCEUTIL.OnArgument (S : String ) ;
begin
    Inherited ;
    tbRazActivite := TToolbarButton97(GetControl('RAZACTIVITE'));
    tbRecCptPieces := TToolbarButton97(GetControl('RECCPTPIECES'));
    tbAffecteDepot := TToolbarButton97(GetControl('AFFECTEDEPOT'));
    Menu := TPopupMenu(GetControl('MENU'));
    mnLance := TMenuItem(GetControl('mnLance'));
    mnInfo := TMenuItem(GetControl('mnInfo'));
    GBInfos := TGroupBox(GetControl('GBINFOS'));
    LstInfos := TListBox(GetControl('LSTINFOS'));
    tbRazActivite.OnCLick := RazClick;
    tbRecCptPieces.OnClick := CalcCptClick;
    tbAffecteDepot.OnClick := AffDepClick;
    mnLance.OnClick := LanceClick;
    mnInfo.OnClick := InfoClick;
end ;

procedure TOF_GCLANCEUTIL.OnClose ;
begin
    Inherited ;
    LstInfos.Free;
end ;
{===============================================================================
 ====================== Choix utilitaire =======================================
 ===============================================================================}
procedure TOF_GCLANCEUTIL.RazClick(Sender : TObject);
begin
    CacheInfo;
    NumChoix := 1;
end;

procedure TOF_GCLANCEUTIL.CalcCptClick(Sender : TObject);
begin
    CacheInfo;
    NumChoix := 2;
end;

procedure TOF_GCLANCEUTIL.AffDepClick(Sender : TObject);
begin
    CacheInfo;
    NumChoix := 3;
end;

{===============================================================================
 ====================== Exécution ==============================================
 ===============================================================================}
procedure TOF_GCLANCEUTIL.CacheInfo;
begin
    GBInfos.Visible := False;
    LstInfos.Clear;
end;

procedure TOF_GCLANCEUTIL.LanceClick(Sender: TObject);
begin
    Lancement(nil);
end;

procedure TOF_GCLANCEUTIL.InfoClick(Sender: TObject);

    procedure AfficheInfo(Titre, Lib : string);
    var TobTmp : TOB;
        Sql : String;
        BlocNote : Variant;
        PosRC : integer;
    begin
        CacheInfo;
        Sql := 'SELECT * FROM JNALEVENT WHERE GEV_TYPEEVENT = "UTI" AND '+
               'GEV_LIBELLE LIKE "'+Lib+'" '+
               'ORDER BY GEV_DATEEVENT DESC';
        TobTmp := TOB.Create('', nil, -1);
        TobTmp.LoadDetailFromSQL(Sql, True, False, 1);
        if TobTmp.Detail.Count > 0 then
        begin
            LstInfos.Items.Add(TraduireMemoire('Exécuté le : ')+
                               DateToStr(TobTmp.Detail[0].GetValue('GEV_DATEEVENT')));
            LstInfos.Items.Add(TraduireMemoire('Par : ')+
                               RechDom('TTUTILISATEUR',TobTmp.Detail[0].GetValue('GEV_UTILISATEUR'),False));
            LstInfos.Items.Add(TraduireMemoire('Description :'));
            BlocNote := TobTmp.Detail[0].GetValue('GEV_BLOCNOTE');
            while BlocNote <> '' do
            begin
                PosRC := pos(#13,BlocNote);
                LstInfos.Items.Add(Copy(BlocNote,1,PosRC-1));
                BlocNote := Copy(BlocNote,PosRC+2,length(BlocNote));
            end;
        end else
        begin
            LstInfos.Items.Add(TraduireMemoire('Cet utilitaire n''a jamais été exécuté.'));
        end;
        TobTmp.Free;
        GBInfos.Caption := TraduireMemoire('Journal sur "')+Titre+'"';
        GBInfos.Visible := True;
    end;

begin
    case Numchoix of
        1 : AfficheInfo(tbRazActivite.caption, TraduireMemoire('Remise à zéro de%'));
        2 : AfficheInfo(tbRecCptPieces.caption, TraduireMemoire('Ajustement des compteurs%'));
        3 : AfficheInfo(tbAffecteDepot.caption, TraduireMemoire('Affectation dépôt%'));
    end;
end;

procedure TOF_GCLANCEUTIL.Lancement(Sender: TObject);
begin
    CacheInfo;
    case Numchoix of
        1 : Assist_RazActivite;
        2 : EntreeMajCompteurSouche;
        3 : Assist_AffecteDepot;
    end;
end;

Initialization
  registerclasses ( [ TOF_GCLANCEUTIL ] ) ;
end.

