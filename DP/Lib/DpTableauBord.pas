//---------------------------------------------------
//--- Auteur : CATALA David
//--- Objet  : Tableaux de bord Comptable / Social
//---------------------------------------------------
unit DpTableauBord;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Vierge, ExtCtrls, HPanel, HSysMenu, HTB97, Grids, Hctrls, uiutil,
{$IFDEF EAGLCLIENT}
  Utob, MenuOlx, UtileAgl,
{$ELSE}
  EdtREtat,Utob,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF} MenuOlg,
{$ENDIF}
  Hmsgbox,HEnt1;

//--------------------------------------------------
//--- Déclaration des constantes
//--------------------------------------------------
const
 COMPTA  =  1;
 PAIE    =  2;

type
  TFDpTableauBord = class(TFVierge)
    Grille: THGrid;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCanResize(Sender: TObject; var NewWidth,
      NewHeight: Integer; var Resize: Boolean);
    procedure BImprimerClick(Sender: TObject);

  private
    { Déclarations privées }
  public
    NMoins_Existe : Boolean;
    TypeAffichage : Integer;
    NumDossier : String;

    TobImprimer : Tob;    

    procedure InitialiserLibelleColonneCompta (Sender: TObject);
    procedure InitialiserLibelleLigneCompta (Sender : TObject);
    procedure InitialiserCelluleCompta (Sender : TObject);
    procedure InitialiserLibelleColonnePaie (Sender: TObject);
    procedure InitialiserLibelleLignePaie (Sender : TObject);
    procedure InitialiserCellulePaie (Sender : TObject);
    procedure GetCellCanvas(ACol,ARow: LongInt;  Canvas : TCanvas; AState : TGridDrawState) ;
    procedure PostDrawCell(ACol,ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);

    function DonnerInformationGrille (UneGrille : THGrid) : Tob;
    procedure ConstruireTobImprimer (UnLibelleLigne : String; NbreChp : Integer; UneTob : Tob);
  end;


//---------------------------------
//--- Déclaration des procédures
//---------------------------------
procedure Aff_TableauBord (NumDossier : String; TypeAffichage : Integer);

implementation

uses DpTableauBordLibrairie, PwdDossier;

{$R *.DFM}

{***********A.G.L.***********************************************
Auteur  ...... : David CATALA
Créé le ...... : 05/06/2003
Modifié le ... :   /  /
Description .. : Affichage de la fiche Tableau de bord
Mots clefs ... : AFF
*****************************************************************}
procedure Aff_TableauBord (NumDossier : String; TypeAffichage : Integer);
Var Fiche : TFDpTableauBord;
    PP : THPanel;
    applisprotec : HTStringList;
begin
 applisprotec := HTStringList.Create;
 // si on n'a pas un password global du dossier...
 if not ToutesApplisProtected(NumDossier, applisprotec) then
   begin
   // ...vérifier qu'on n'a pas un password par appli
   if ( (TypeAffichage=COMPTA) and (applisprotec.IndexOf('CCS5.EXE')>-1) )
   or ( (TypeAffichage=PAIE)   and (applisprotec.IndexOf('CPS5.EXE')>-1) ) then
     if Not VerifPwdDossier(NumDossier, False) then
       begin applisprotec.Free; FMenuG.VireInside(Nil); exit; end;
   end;
 applisprotec.Free;

 Fiche:=TFDpTableauBord.Create (Application);
 Fiche.TypeAffichage:=TypeAffichage;
 Fiche.NumDossier:=NumDossier;
 PP:=FindInsidePanel;
 if PP=nil then
  begin
   try
    Fiche.showModal;
   finally;
    Fiche.Free;
   end;
  end
 else
  begin
   InitInside (Fiche,PP);
   Fiche.Show;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : David CATALA
Créé le ...... : 12/06/2003
Modifié le ... :   /  /
Description .. : Création de la fiche
Mots clefs ... :
*****************************************************************}
procedure TFDpTableauBord.FormCreate(Sender: TObject);
begin
 Inherited;
 Grille.Ctl3D:=False;
 Grille.GetCellCanvas:=GetCellCanvas ;
 Grille.PostDrawCell:=PostDrawCell;
end;

{***********A.G.L.***********************************************
Auteur  ...... : David CATALA
Créé le ...... : 12/06/2003
Modifié le ... :   /  /
Description .. : Affichage de la fiche
Mots clefs ... :
*****************************************************************}
procedure TFDpTableauBord.FormShow(Sender: TObject);
begin
 inherited;
 if (TypeAffichage=COMPTA) then
  begin
   Self.Caption:='Tableau de bord comptable';
   Grille.EnabledBlueButton:=True;
   InitialiserLibelleColonneCompta (Sender);
   InitialiserLibelleLigneCompta (Sender);
   InitialiserCelluleCompta (Sender);
  end
 else
  begin
   Self.Caption:='Tableau de bord social';
   Grille.EnabledBlueButton:=True;
   InitialiserLibelleColonnePaie (Sender);
   InitialiserLibelleLignePaie (Sender);
   InitialiserCellulePaie (Sender);
  end;
 UpdateCaption(Self);
end;

{***********A.G.L.***********************************************
Auteur  ...... : David CATALA
Créé le ...... : 05/06/2003
Modifié le ... :   /  /
Description .. : Définit les colonnes de la grille Compta
Mots clefs ... :
*****************************************************************}
Procedure TFDpTableauBord.InitialiserLibelleColonneCompta (Sender: TObject);
var SSql       : String;
    QDos       : Tquery;
    NumColonne : Integer;
    Indice     : Integer;
begin
 Grille.ColCount:=4;
 Grille.ColWidths [0]:=200;

 SSql:='SELECT * from DPTABCOMPTA where DTC_NODOSSIER="'+NumDossier+'"';
  try
   QDos:=Opensql (SSql, TRUE);
   for Indice:=1 to Grille.ColCount-1 do
    begin
     Grille.ColWidths [Indice]:=(Grille.Width-Grille.ColWidths [0]) div (Grille.ColCount-1);
     if (not QDos.Eof) then
      begin
       //--- Determine le numéro de la colonne
       if (QDos.FindField ('DTC_MILLESIME').AsString='N-') then
        NumColonne:=1
       else
        if (QDos.FindField ('DTC_MILLESIME').AsString='N') then
         NumColonne:=2
        else
         NumColonne:=3;

       //--- Affichage des libelles colonnes
       Grille.Cells [NumColonne,0]:=QDos.FindField ('DTC_LIBEXERCICE').AsString;
       QDos.Next;
      end;
    end;
   Ferme (QDos);
  except
   On E:Exception do PGIInfo(E.Message, 'Lecture DPTABCOMPTA');
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : David CATALA
Créé le ...... : 11/06/2003
Modifié le ... :   /  /
Description .. : Défini les lignes de la grille Compta
Mots clefs ... :
*****************************************************************}
procedure TFDpTableauBord.InitialiserLibelleLigneCompta (Sender : TObject);
begin
 Grille.RowCount:=23;
 Grille.Cells [0,1]:='Durée (en mois)';
 Grille.Cells [0,2]:='Date début';
 Grille.Cells [0,3]:='Date fin';

 Grille.Cells [0,4]:='Données économiques';
 Grille.Cells [0,5]:='Chiffre d''affaires';
 Grille.Cells [0,6]:='Marge totale';
 Grille.Cells [0,7]:='Valeur ajoutée';
 Grille.Cells [0,8]:='Excédent brut d''exploitation';
 Grille.Cells [0,9]:='Résultat d''exploitation';
 Grille.Cells [0,10]:='Résultat courant avant impôt';
 Grille.Cells [0,11]:='Resultat de l''exercice';

 Grille.Cells [0,12]:='Avancement du dossier';
 Grille.Cells [0,13]:='Nombre d''ecriture(s)';
 Grille.Cells [0,14]:='En-cours de saisie';
 Grille.Cells [0,15]:='Dernière saisie effectuée au';
 Grille.Cells [0,16]:='Sur le journal';
 Grille.Cells [0,17]:='Par';
 Grille.Cells [0,18]:='Le';

 Grille.Cells [0,19]:='Immobilisations';
 Grille.Cells [0,20]:='Nombre d''immobilisations';
 Grille.Cells [0,21]:='Entrées de l''exercice';
 Grille.Cells [0,22]:='Sorties de l''exercice';

end;

{***********A.G.L.***********************************************
Auteur  ...... : David CATALA
Créé le ...... : 12/06/2003
Modifié le ... :   /  /
Description .. : Définit le contenu des cellules de la grille Compta
Mots clefs ... :
*****************************************************************}
procedure TFDpTableauBord.InitialiserCelluleCompta (Sender : TObject);
var SSql       : String;
    QDos       : Tquery;
    NumColonne : Integer;
    Indice     : Integer;
begin
 NMoins_Existe:=False;
 SSql:='SELECT * from DPTABCOMPTA where DTC_NODOSSIER="'+NumDossier+'"';
  try
   QDos:=Opensql (SSql, TRUE);

   for Indice:=1 to Grille.ColCount-1 do
    begin
     if (not QDos.Eof) then
      begin
       //--- Détermine le numéro de la colonne
       if (QDos.FindField ('DTC_MILLESIME').AsString='N-') then
        begin
         NMoins_Existe:=True;
         NumColonne:=1
        end
       else
        if (QDos.FindField ('DTC_MILLESIME').AsString='N') then
         NumColonne:=2
        else
         NumColonne:=3;

       //--- Affichage des valeurs des cellules
       Grille.Cells [NumColonne,1]:=QDos.FindField ('DTC_DUREE').AsString;
       Grille.Cells [NumColonne,2]:=QDos.FindField ('DTC_DATEDEB').AsString;
       Grille.Cells [NumColonne,3]:=QDos.FindField ('DTC_DATEFIN').AsString;
       Grille.Cells [NumColonne,5]:=StrFMontant (Valeur (QDos.FindField ('DTC_CA').AsString),15,2,'',True);//+' '+GetParamSoc ('So_DevicePrinc');;
       Grille.Cells [NumColonne,6]:=StrFMontant (Valeur (QDos.FindField ('DTC_MARGETOTALE').AsString),15,2,'',True);
       Grille.Cells [NumColonne,7]:=StrFMontant (Valeur (QDos.FindField ('DTC_VALEURAJOUTEE').AsString),15,2,'',True);
       Grille.Cells [NumColonne,8]:=StrFMontant (Valeur (QDos.FindField ('DTC_EXCEDBRUT').AsString),15,2,'',True);
       Grille.Cells [NumColonne,9]:=StrFMontant (Valeur (QDos.FindField ('DTC_RESULTEXPLOIT').AsString),15,2,'',True);
       Grille.Cells [NumColonne,10]:=StrFMontant (Valeur (QDos.FindField ('DTC_RESULTCOURANT').AsString),15,2,'',True);
       Grille.Cells [NumColonne,11]:=StrFMontant (Valeur (QDos.FindField ('DTC_RESULTEXERC').AsString),15,2,'',True);
       if (NumColonne<>1) then
        begin
         Grille.Cells [NumColonne,13]:=FormaterNombre (QDos.FindField ('DTC_NBECRITURE').AsString);
         Grille.Cells [NumColonne,14]:=QDos.FindField ('DTC_ENCOURSSAISIE').AsString;
         Grille.Cells [NumColonne,15]:=QDos.FindField ('DTC_DERNSAISIE').AsString;
         Grille.Cells [NumColonne,16]:=QDos.FindField ('DTC_DERNJOURNAL').AsString;
         Grille.Cells [NumColonne,17]:=RechDom ('TTUTILISATEUR',QDos.FindField ('DTC_UTILISATSAISIE').AsString,False);
         Grille.Cells [NumColonne,18]:=QDos.FindField ('DTC_DATESAISIE').AsString;
        end;
       Grille.Cells [NumColonne,20]:=FormaterNombre (QDos.FindField ('DTC_NBLIGNEIMMO').AsString);
       Grille.Cells [NumColonne,21]:=FormaterNombre (QDos.FindField ('DTC_NBENTREEIMMO').AsString);
       Grille.Cells [NumColonne,22]:=FormaterNombre (QDos.FindField ('DTC_NBSORTIEIMMO').AsString);
       QDos.Next;
      end;
    end;

   Ferme (QDos);
  except
   On E:Exception do PGIInfo(E.Message, 'Lecture DPTABCOMPTA');
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : David CATALA
Créé le ...... : 05/06/2003
Modifié le ... :   /  /
Description .. : Définit les colonnes de la grille Paie
Mots clefs ... :
*****************************************************************}
Procedure TFDpTableauBord.InitialiserLibelleColonnePaie (Sender: TObject);
var SSql       : String;
    QDos       : Tquery;
    NumColonne : Integer;
    Indice     : Integer;
begin
 Grille.ColCount:=3;
 Grille.ColWidths [0]:=200;
 NumColonne:=1;

 SSql:='SELECT * from DPTABPAIE where DTP_NODOSSIER="'+NumDossier+'"';
  try
   QDos:=Opensql (SSql, TRUE);
   for Indice:=1 to Grille.ColCount-1 do
    begin
     Grille.ColWidths [Indice]:=(Grille.Width-Grille.ColWidths [0]) div (Grille.ColCount-1);
     if (not QDos.Eof) then
      begin
       //--- Détermine le numéro de la colonne
       if (QDos.FindField ('DTP_MILLESIME').AsString='N-') then
        begin
         NMoins_Existe:=True;
         NumColonne:=1
        end
       else
        if (QDos.FindField ('DTP_MILLESIME').AsString='N') then
         NumColonne:=2;

       if (not QDos.Eof) then
        Grille.Cells [NumColonne,0]:=QDos.FindField ('DTP_LIBEXERCICE').AsString;
      end;
     QDos.Next;
    end;
   Ferme (QDos);
  except
   On E:Exception do PGIInfo(E.Message, 'Lecture DPTABPAIE');
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : David CATALA
Créé le ...... : 11/06/2003
Modifié le ... :   /  /
Description .. : Défini les lignes de la grille Paie
Mots clefs ... :
*****************************************************************}
procedure TFDpTableauBord.InitialiserLibelleLignePaie (Sender : TObject);
begin
 Grille.RowCount:=9;
 Grille.Cells [0,1]:='Date début';
 Grille.Cells [0,2]:='Date fin';
 Grille.Cells [0,3]:='Décalage paie';
 Grille.Cells [0,4]:='Montant DADS';
 Grille.Cells [0,5]:='Effectif fin d''exercice';
 Grille.Cells [0,6]:='Entrée(s) dans l''année';
 Grille.Cells [0,7]:='Sortie(s) dans l''année';
 Grille.Cells [0,8]:='Mode réglement salaire';
end;

{***********A.G.L.***********************************************
Auteur  ...... : David CATALA
Créé le ...... : 12/06/2003
Modifié le ... :   /  /
Description .. : Définit le contenu des cellules de la grille Paie
Mots clefs ... :
*****************************************************************}
procedure TFDpTableauBord.InitialiserCellulePaie (Sender : TObject);
var SSql       : String;
    QDos       : Tquery;
    NumColonne : Integer;
    Indice     : Integer;
begin
 NumColonne:=1;
 NMoins_Existe:=False;
 SSql:='SELECT * from DPTABPAIE where DTP_NODOSSIER="'+NumDossier+'"';
  try
   QDos:=Opensql (SSql, TRUE);
   for Indice:=1 to Grille.ColCount-1 do
    begin
     if (not QDos.Eof) then
      begin
       //--- Détermine le numéro de la colonne
       if (QDos.FindField ('DTP_MILLESIME').AsString='N-') then
        begin
         NMoins_Existe:=True;
         NumColonne:=1
        end
       else
        if (QDos.FindField ('DTP_MILLESIME').AsString='N') then
         NumColonne:=2;

       //--- Affichage des valeurs des cellules
       Grille.Cells [NumColonne,1]:=QDos.FindField ('DTP_DATEDEB').AsString;
       Grille.Cells [NumColonne,2]:=QDos.FindField ('DTP_DATEFIN').AsString;
       Grille.Cells [NumColonne,3]:=QDos.FindField ('DTP_DECALAGEPAIE').AsString;
       Grille.Cells [NumColonne,4]:=StrFMontant (Valeur (QDos.FindField ('DTP_MONTANTDADS').AsString),15,2,'',True);
       Grille.Cells [NumColonne,5]:=QDos.FindField ('DTP_EFFECTIF').AsString;
       Grille.Cells [NumColonne,6]:=QDos.FindField ('DTP_NBENTREES').AsString;
       Grille.Cells [NumColonne,7]:=QDos.FindField ('DTP_NBSORTIES').AsString;
       Grille.Cells [NumColonne,8]:=RechDom ('PGMODEREGLE',QDos.FindField ('DTP_MODEREGL').AsString,False);
      end;
     QDos.Next;
    end;
   Ferme (QDos);
  except
   On E:Exception do PGIInfo(E.Message, 'Lecture DPTABPAIE');
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : David CATALA
Créé le ...... : 11/06/2003
Modifié le ... :   /  /
Description .. : Définit le contenu de la cellule aprés que la cellule soit
Suite ........ : dessiné par le canvas
Mots clefs ... :
*****************************************************************}
procedure TFDpTableauBord.GetCellCanvas(ACol,ARow: LongInt;  Canvas : TCanvas; AState : TGridDrawState) ;
begin
 if (TypeAffichage=COMPTA) then
  if (ACol=0) and ((ARow=4) or (ARow=12) or (ARow=19)) then
   Grille.Canvas.Font.Style:=[FsBold];
end;

{***********A.G.L.***********************************************
Auteur  ...... : David CATALA
Créé le ...... : 11/06/2003
Modifié le ... :   /  /
Description .. : Définit le contenu de la cellule avant que la cellule soit
Suite ........ : dessiné par le canvas
Mots clefs ... :
*****************************************************************}
procedure TFDpTableauBord.PostDrawCell(ACol,ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin
 if (TypeAffichage=COMPTA) then
  begin
   //--- Alignement du Texte
   if (ACol>0) and (ARow>0) and (ARow<5) then
    EcrireTexte (CENTRE,ACol,ARow,Grille)
   else
    if  (ACol>0) and (ARow>12) and (ARow<19) then
     EcrireTexte (CENTRE,ACol,ARow,Grille)
    else
     if  (ACol>0) and (ARow>4) and (ARow<12) then
      EcrireTexte (DROITE,ACol,ARow,Grille)
     else
      if  (ACol>0) and (ARow>19) and (ARow<23) then
       EcrireTexte (CENTRE,ACol,ARow,Grille);

   if ((ACol=1) or (ACol=2) or (ACol=3)) and ((ARow=4) or (ARow=12) or (ARow=19)) then
    NoirsirGrille (ACol,ARow,Grille);

   if (ACol=1) and ((ARow=14) or (ARow=15) or (ARow=16) or (ARow=17) or (ARow=18)) then
    GriserGrille (Acol,ARow,Grille);

   if (not NMoins_Existe) and (ACol=1) then
    GriserGrille (Acol,ARow,Grille);
  end
 else
  begin
   //--- Alignement du Texte
   if (ACol>0) and (ARow=4) then
    EcrireTexte (DROITE,ACol,ARow,Grille)
   else
    if (ACol>0) then
     EcrireTexte (CENTRE,ACol,ARow,Grille);

   //--- Grisement des cellules
   if (not NMoins_Existe) and (ACol=1) then
    GriserGrille (Acol,ARow,Grille);
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : David CATALA
Créé le ...... : 11/06/2003
Modifié le ... :   /  /
Description .. : Réinitialise la grille sur le redimensionnement
                 de la fenêtre
Mots clefs ... :
*****************************************************************}
procedure TFDpTableauBord.FormCanResize(Sender: TObject; var NewWidth,
  NewHeight: Integer; var Resize: Boolean);
begin
  inherited;
  if (TypeAffichage=COMPTA) then
   begin
    InitialiserLibelleColonneCompta (Sender);
    InitialiserLibelleLigneCompta (Sender);
    InitialiserCelluleCompta (Sender);
   end
  else
   begin
    InitialiserLibelleColonnePaie (Sender);
    InitialiserLibelleLignePaie (Sender);
    InitialiserCellulePaie (Sender);
   end;
end;


//------------------------------------------------------------------------------
//--- Nom   : BImprimerClick
//------------------------------------------------------------------------------
procedure TFDpTableauBord.BImprimerClick(Sender: TObject);
var UneTob : Tob;
    Indice : Integer;
begin
 TobImprimer:=TOB.Create('',nil,-1);

 //--- Récupération des informations contenu dans  la grille
 UneTob:=DonnerInformationGrille (Grille);
 if (UneTob.Detail.Count>0) then
  begin
   ConstruireTobImprimer ('ENTETE',Grille.ColCount,UneTob.Detail [0]);
   for Indice:=1 to UneTob.Detail.count-1 do
    ConstruireTobImprimer ('DETAIL',Grille.ColCount,UneTob.Detail [Indice]);
  end;
 FreeAndNil (uneTob);

 //--- Impression de l'état Tableau de bord Compta / Paie
 if TobImprimer.Detail.Count > 0 then
  begin
   if (TypeAffichage=COMPTA) then
    LanceEtatTob( 'E','DPG','TB2',TobImprimer,True,False,False,nil,'','Edition Tableau de bord Compta',False)
   else
    LanceEtatTob( 'E','DPG','TB3',TobImprimer,True,False,False,nil,'','Edition Tableau de bord Paie',False)
  end
 else
  PgiInfo ('Aucune donnée à imprimer.',TitreHalley);
 FreeAndNil (TobImprimer);
end;

//------------------------------------------------------------------------------
//--- Nom   : DonnerInformationGrille
//------------------------------------------------------------------------------
function TFDpTableauBord.DonnerInformationGrille (UneGrille : THGrid) : Tob;
Var TobGrille,UneTobEnreg : Tob;
    NomChp,ValChp         : String;
    IndiceCol,IndiceLigne : Integer;
begin
 TobGrille:=Tob.Create ('',nil,-1);

 for IndiceLigne:=0 to UneGrille.RowCount-1 do
  begin
   UneTobEnreg:=TOB.Create ('',TobGrille,-1);
   for IndiceCol:=0 to UneGrille.ColCount-1 do
    begin
     NomChp:='CHP'+Format ('%2.2d',[IndiceCol]);
     ValChp:=UneGrille.Cells [IndiceCol,IndiceLigne];
     UneTobEnreg.AddChampSupValeur (NomChp,ValChp);
    end;
  end;

 Result:=TobGrille;
end;

//------------------------------------------------------------------------------
//--- Nom   : ConstruireTobImprimer
//------------------------------------------------------------------------------
procedure TFDpTableauBord.ConstruireTobImprimer (UnLibelleLigne : String; NbreChp : Integer; UneTob : Tob);
var Indice      : Integer;
    TbChpVal    : String;
    UneTobEnreg : Tob;
begin

 UneTobEnreg:=TOB.Create ('',TobImprimer,-1);
 for Indice:=0 to 14 do
  UneTobEnreg.AddChampSupValeur ('TBChp'+intToStr (Indice),'');

 UneTobEnreg.AddChampSupValeur ('TBChp0',UnLibelleLigne);
 for Indice:=1 to NbreChp do
  begin
   TbChpVal:=UneTob.Getvalue ('CHP'+Format ('%2.2d',[Indice-1]));
   if (TbChpVal<>'') then
    UneTobEnreg.PutValue ('TBChp'+IntToStr (Indice),TbChpVal)
   else
    UneTobEnreg.PutValue ('TBChp'+IntToStr (Indice),' ');
  end;
end;

end.
