{***********UNITE*************************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 22/02/2007
Modifié le ... :   /  /
Description .. : Gestion spécifique des cumuls alimentant les rubriques
Suite ........ : de cotisations et de rémunérations
Suite ........ : Source TOF de la FICHE : PGCUMULGESTIONSPEC ()
Mots clefs ... : TOF;PGCUMULGESTIONSPEC;PAIE;GESTIONSPECIFIQUE
*****************************************************************}
{
PT1    : 31/05/2007 VG V_72 Modification du message - FQ N°14290
PT2-1  : 05/06/2007 VG V_72 Traitement du cas où la rubrique ne possède pas de
                            gestion associée - FQ N°14302
PT2-2  : 05/06/2007 VG V_72 Optimisation
PT2-3  : 05/06/2007 VG V_72 Sauvegarde avant impression
PT3    : 16/07/2007 VG V_72 Affichage dans une seule grille
PT4    : 05/12/2007 VG V_72 Validation ne marche pas en CWAS - FQ N°15012
}
Unit UTOF_PGCUMULGESTIONSPEC;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
     FE_Main,
     EdtREtat,
{$else}
     eMul,
     MaineAgl,
     UtileAgl,
{$ENDIF}
     UTob,
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     PgOutils,
     PgOutils2,
     HTB97,
     uTobDebug,
     Vierge;

Type
  TOF_PGCUMULGESTIONSPEC = Class (TOF)
  public
    procedure OnNew                    ; override;
    procedure OnDelete                 ; override;
    procedure OnUpdate                 ; override;
    procedure OnLoad                   ; override;
    procedure OnArgument (S : String ) ; override;
    procedure OnDisplay                ; override;
    procedure OnClose                  ; override;
    procedure OnCancel                 ; override;
  private {Déclarations privées}
    TOB_Cumuls, TOB_CumulsAutorises, TOB_CURUB, TOB_CURUB_Herit : TOB;
    Cumuls : THGrid;
    CEG, STD, DOS, LectureSeule, Modifie, Retour: Boolean;
    Quoi: string;
    Libelle, Nature, Rubrique, TypeAction : string;
    procedure AfficheGrid;
    procedure Sauve;
    procedure CumulsDblClick (Sender: TObject);
    procedure CumulsCellExit (Sender: TObject; var ACol,ARow: Integer;
                              var Cancel: Boolean);
    procedure Impression(Sender: TObject);
  end;

Implementation

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 22/02/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGCUMULGESTIONSPEC.OnNew;
begin
Inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 22/02/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGCUMULGESTIONSPEC.OnDelete;
begin
Inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 22/02/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGCUMULGESTIONSPEC.OnUpdate;
begin
Inherited;
Sauve;
SetControlProperty ('BVALIDER', 'ModalResult', mrOK);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 22/02/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGCUMULGESTIONSPEC.OnLoad;
begin
Inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 22/02/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGCUMULGESTIONSPEC.OnArgument (S : String);
var
TypeP, St : string;
begin
Inherited;
Nature:= Trim (ReadTokenPipe (S, ';'));
Rubrique:= Trim (ReadTokenPipe (S, ';'));
Libelle:= Trim (ReadTokenPipe (S, ';'));
TypeAction:= Trim (ReadTokenPipe (S, ';'));

if (Quoi='C') then
   TypeP:= 'CEG'
else
if (Quoi='S') then
   TypeP:= 'STD'
else
   TypeP:= 'DOS';

AccesPredefini (TypeP, CEG, STD, DOS);

if (TypeP = 'CEG') then
   LectureSeule:= (CEG=False)
else
if (TypeP = 'STD') then
   LectureSeule:= (STD=False)
else
   LectureSeule:= (DOS=False);
if LectureSeule then
   begin
{PT3
   SetControlEnabled ('BNONAFFECTE_GAIN', FALSE);
   SetControlEnabled ('BGAIN_NONAFFECTE', FALSE);
   SetControlEnabled ('BNONAFFECTE_RETENU', FALSE);
   SetControlEnabled ('BRETENU_NONAFFECTE', FALSE);
   SetControlEnabled ('BGAIN_RETENU', FALSE);
   SetControlEnabled ('BRETENU_GAIN', FALSE);
}
   SetControlEnabled ('GRID_CUMULS', FALSE);
//FIN PT3
   SetControlEnabled ('BVALIDER', FALSE);
   end;

SetControlText ('LRUBRIQUE', Rubrique);
SetControlText ('LLIBELLE', Libelle);
SetControlText ('LNATURERUB', Nature);
Modifie:= FALSE;
Retour:= FALSE;
{$IFDEF EAGLCLIENT}
AvertirCacheServer ('CUMULPAIE');
{$ENDIF}

// CHARGEMENT DE LA TABLE DES CUMULS
TOB_Cumuls:= TOB.Create ('Table des cumuls', nil, -1);
TOB_Cumuls.LoadDetailDB ('CUMULPAIE', '', '', nil, False);
{PT3
TOB_CumulsAffiches:= TOB.Create ('Table2 des cumuls', nil, -1);
}

TOB_CumulsAutorises:= TOB.Create ('Table des cumuls autorisés', nil, -1);
St:= 'SELECT *'+
     ' FROM PGEXCEPTIONS WHERE'+
     ' PEN_NATURERUB="CUM"';
{Optimisation - PT2-2
Q:= OpenSql (St, TRUE);
TOB_CumulsAutorises.LoadDetailDB ('PGEXCEPTIONS', '', '', Q, False);
Ferme (Q);
}
TOB_CumulsAutorises.LoadDetailDBFromSQL ('PGEXCEPTIONS', St);
//FIN PT2-2
{PT3
TOB_CumulCache:= TOB.Create ('Table des cumuls cachés', nil, -1);
}

// CHARGEMENT DE LA TABLE DES CUMULS / RUBRIQUES
TOB_CURUB:= TOB.Create ('Tables des cumuls par rubrique', nil, -1);
St:= 'SELECT *'+
     ' FROM CUMULRUBDOSSIER WHERE'+
     ' PKC_NATURERUB="'+Nature+'" AND'+
     ' PKC_RUBRIQUE="'+Rubrique+'"';
{Optimisation - PT2-2
Q:= OpenSql (St, TRUE);
TOB_CURUB.LoadDetailDB ('CUMULRUBDOSSIER', '', '', Q, FALSE);
Ferme (Q);
}
TOB_CURUB.LoadDetailDBFromSQL ('CUMULRUBDOSSIER', St);
//FIN PT2-2

TOB_CURUB_Herit := TOB.Create('Tables des cumuls par rubrique', nil, -1);
St:= 'SELECT *'+
     ' FROM CUMULRUBRIQUE WHERE'+
     ' PCR_NATURERUB="'+Nature+'" AND'+
     ' ##PCR_PREDEFINI## PCR_RUBRIQUE="'+Rubrique+'"';
{Optimisation - PT2-2
Q:= OpenSql(st, TRUE);
TOB_CURUB_Herit.LoadDetailDB ('CUMULRUBRIQUE', '', '', Q, FALSE);
Ferme(Q);
}
TOB_CURUB_Herit.LoadDetailDBFromSQL ('CUMULRUBRIQUE', St);
//FIN PT2-2

{PT3
TOB_AncienCURUB:= TOB.Create ('tables ANCIEN', nil, -1);
TOB_AncienCURUB.Dupliquer (TOB_CURUB, TRUE, TRUE);
TOB_AncienCURUB_Herit:= TOB.Create ('tables ANCIEN', nil, -1);
TOB_AncienCURUB_Herit.Dupliquer (TOB_CURUB_Herit, TRUE, TRUE);

// CREATION DE LA TOB DES CUMULS AFFECTES EN GAIN
TOB_CUGain:= TOB.Create ('liste gain', nil, -1);
TOB_CUGain_Herit:= TOB.Create ('liste gain', nil, -1);

// CREATION DE LA TOB DES CUMULS AFFECTES EN RETENU
TOB_CURetenu:= TOB.Create ('liste retenu', nil, -1);
TOB_CURetenu_Herit:= TOB.Create ('liste retenu', nil, -1);

VireDejaUtilise;
}

Cumuls:= THGRID (GetControl ('GRID_CUMULS'));
if (Cumuls <> nil) then
   begin
   Cumuls.OnDblClick:= CumulsDblClick;
{PT3
   Cumuls.OnDragDrop:= Depose_Objet;
   Cumuls.OnDragOver:= Test_Depose_Objet;
   Cumuls.OnMouseDown:= CumulsMouseDown;
   Cumuls.ColAligns [0]:= taCenter;
   Cumuls.ColWidths [0]:= 26;
   Cumuls.ColAligns [1]:= taLeftJustify;
   Cumuls.ColWidths [1]:= 158;
   Cumuls.ColAligns [2]:= taCenter;
   Cumuls.ColWidths [2]:= 26;
   Cumuls.ColTypes  [2]:= 'B';
   Cumuls.ColFormats[2]:= IntToStr(Ord(csCoche));
}
   Cumuls.OnCellExit:= CumulsCellExit;
   Cumuls.ColAligns   [0]:= taCenter;
   Cumuls.ColWidths   [0]:= 32;
   Cumuls.ColEditables[0]:= False;
   Cumuls.ColAligns   [1]:= taLeftJustify;
   Cumuls.ColWidths   [1]:= 256;
   Cumuls.ColEditables[1]:= False;
   Cumuls.ColAligns   [2]:= taCenter;
   Cumuls.ColWidths   [2]:= 64;
   Cumuls.ColEditables[2]:= False;
   Cumuls.ColTypes    [2]:= 'B';
   Cumuls.ColFormats  [2]:= IntToStr(Ord(csCoche));
   Cumuls.ColAligns   [3]:= taCenter;
   Cumuls.ColWidths   [3]:= 64;
   Cumuls.ColFormats  [3]:= 'CB=PGSENSRUBRIQUE';
//FIN PT3
   end;

{PT3
Gain:= THGRID (GetControl ('GRID_GAIN'));
if (Gain <> nil) then
   begin
   Gain.OnDblClick:= GainDblClick;
   Gain.OnDragDrop:= Depose_Objet;
   Gain.OnDragOver:= Test_Depose_Objet;
   Gain.OnMouseDown:= CumulsMouseDown;
   Gain.ColAligns [0]:= taCenter;
   Gain.ColWidths [0]:= 26;
   Gain.ColAligns [1]:= taLeftJustify;
   Gain.ColWidths [1]:= 178;
   Gain.ColAligns [2]:= taCenter;
   Gain.ColWidths [2]:= 38;
   Gain.ColFormats[2]:= 'CB=PGSENSMVTCP';
   end;

Gain_Herit:= THGRID (GetControl ('GRID_GAIN_HERIT'));
if (Gain <> nil) then
   begin
   Gain_Herit.OnDblClick:= GainDblClick;
   Gain_Herit.OnDragDrop:= Depose_Objet;
   Gain_Herit.OnMouseDown:= CumulsMouseDown;
   Gain_Herit.ColAligns [0]:= taCenter;
   Gain_Herit.ColWidths [0]:= 26;
   Gain_Herit.ColAligns [1]:= taLeftJustify;
   Gain_Herit.ColWidths [1]:= 158;
   Gain_Herit.ColAligns [2]:= taCenter;
   Gain_Herit.ColWidths [2]:= 26;
   Gain_Herit.ColTypes  [2]:= 'B';
   Gain_Herit.ColFormats[2]:= IntToStr(Ord(csCoche));
   end;

Retenu:= THGRID (GetControl ('GRID_RETENU'));
if (Retenu <> nil) then
   begin
   Retenu.OnDblClick:= RetenuDblClick;
   Retenu.OnDragDrop:= Depose_Objet;
   Retenu.OnDragOver:= Test_Depose_Objet;
   Retenu.OnMouseDown:= CumulsMouseDown;
   Retenu.ColAligns [0]:= taCenter;
   Retenu.ColWidths [0]:= 26;
   Retenu.ColAligns [1]:= taLeftJustify;
   Retenu.ColWidths [1]:= 178;
   Retenu.ColAligns [2]:= taCenter;
   Retenu.ColWidths [2]:= 38;
   Retenu.ColFormats[2]:= 'CB=PGSENSMVTCP';
   end;

Retenu_Herit:= THGRID (GetControl ('GRID_RETENU_HERIT'));
if (Retenu_Herit <> nil) then
   begin
   Retenu_Herit.OnDblClick:= RetenuDblClick;
   Retenu_Herit.OnDragDrop:= Depose_Objet;
   Retenu_Herit.OnMouseDown:= CumulsMouseDown;
   Retenu_Herit.ColAligns [0]:= taCenter;
   Retenu_Herit.ColWidths [0]:= 26;
   Retenu_Herit.ColAligns [1]:= taLeftJustify;
   Retenu_Herit.ColWidths [1]:= 158;
   Retenu_Herit.ColAligns [2]:= taCenter;
   Retenu_Herit.ColWidths [2]:= 26;
   Retenu_Herit.ColTypes  [2]:= 'B';
   Retenu_Herit.ColFormats[2]:= IntToStr(Ord(csCoche));
   end;

BNonAffecte_Gain:= TToolbarButton97 (GetControl ('BNONAFFECTE_GAIN'));
if (BNonAffecte_Gain <> nil) then
   BNonAffecte_Gain.OnClick:= BNonAffecte_GainClick;

BNonAffecte_Retenu:= TToolbarButton97 (GetControl ('BNONAFFECTE_RETENU'));
if (BNonAffecte_Retenu <> nil) then
   BNonAffecte_Retenu.OnClick:= BNonAffecte_RetenuClick;

BGain_NonAffecte:= TToolbarButton97 (GetControl ('BGAIN_NONAFFECTE'));
if (BGain_NonAffecte <> nil) then
   BGain_NonAffecte.OnClick:= BGain_NonAffecteClick;

BRetenu_NonAffecte:= TToolbarButton97 (GetControl ('BRETENU_NONAFFECTE'));
if (BRetenu_NonAffecte <> nil) then
   BRetenu_NonAffecte.OnClick:= BRetenu_NonAffecteClick;

BGain_Retenu:= TToolbarButton97 (GetControl ('BGAIN_RETENU'));
if (BGain_Retenu <> nil) then
   BGain_Retenu.OnClick:= BGain_RetenuClick;

BRetenu_Gain:= TToolbarButton97 (GetControl ('BRETENU_GAIN'));
if (BRetenu_Gain <> nil) then
   BRetenu_Gain.OnClick:= BRetenu_GainClick;
}

TFVierge(Ecran).BImprimer.OnClick:= Impression;

AfficheGrid;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 22/02/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGCUMULGESTIONSPEC.OnClose;
begin
Inherited;
FreeAndNil (TOB_Cumuls);
{PT3
FreeAndNil (TOB_CumulsAffiches);
}
FreeAndNil (TOB_CumulsAutorises);
{PT3
FreeAndNil (TOB_CumulCache);
}
FreeAndNil (TOB_CURUB);
{PT3
FreeAndNil (TOB_AncienCURUB);
FreeAndNil (TOB_CUGain);
FreeAndNil (TOB_CURetenu);
}
FreeAndNil (TOB_CURUB_Herit);
{PT3
FreeAndNil (TOB_AncienCURUB_Herit);
FreeAndNil (TOB_CUGain_Herit);
FreeAndNil (TOB_CURetenu_Herit);
}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 22/02/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGCUMULGESTIONSPEC.OnDisplay ();
begin
Inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 22/02/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGCUMULGESTIONSPEC.OnCancel ();
begin
Inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 22/02/2007
Modifié le ... :   /  /
Description .. : Affichage du contenu des grilles
Mots clefs ... : TOF;PGCUMULGESTIONSPEC;PAIE;GESTIONSPECIFIQUE
*****************************************************************}
procedure TOF_PGCUMULGESTIONSPEC.AfficheGrid;
var
i : integer;
TOB_CumulsD, TOBDetail : TOB;
Sens : string;
begin
{PT3
TOB_CumulsAffiches.Dupliquer (TOB_Cumuls, True, True);
//TobDebug (TOB_CumulsAffiches);

For i := 0 to TOB_Cumuls.Detail.Count-1 do
    begin
    TOB_CumulsD:= TOB_Cumuls.Detail[i];
    TOB_CumulsAutorisesD:= TOB_CumulsAutorises.FindFirst (['PEN_RUBRIQUE'],
                                                          [TOB_CumulsD.GetValue('PCL_CUMULPAIE')],
                                                          True);
    TOB_CumulsAffichesD:= TOB_CumulsAffiches.FindFirst (['PCL_CUMULPAIE'],
                                                        [TOB_CumulsD.GetValue ('PCL_CUMULPAIE')],
                                                        True);

    if (TOB_CumulsD.FieldExists('AUTOR')) then
       begin
       TOB_CumulsD.PutValue ('AUTOR', TOB_CumulsAutorisesD<>nil);
       TOB_CumulsAffichesD.PutValue ('AUTOR', TOB_CumulsAutorisesD<>nil);
       end
    else
       begin
       TOB_CumulsD.AddChampSupValeur ('AUTOR', TOB_CumulsAutorisesD<>nil);
       TOB_CumulsAffichesD.AddChampSupValeur ('AUTOR', TOB_CumulsAutorisesD<>nil);
       end;
    if ((TOB_CumulsD.GetValue ('AUTOR')=False) and
       (TOB_CumulsAffichesD <> nil)) then
       FreeAndNil (TOB_CumulsAffichesD);
    end;
//TobDebug (TOB_CumulsAffiches);
TOB_CumulsAffiches.Detail.Sort ('PCL_CUMULPAIE');
TOB_CumulsAffiches.PutGridDetail (Cumuls, FALSE, TRUE,
                                  'PCL_CUMULPAIE;PCL_LIBELLE;AUTOR', TRUE);
//Cumuls.SortGrid (0, False);

TOB_CUGain.Detail.Sort ('PKC_CUMULPAIE');
TOB_CUGain.PutGridDetail (Gain, FALSE, TRUE,
                          'PKC_CUMULPAIE;PKC_LIBELLE;PKC_PLUSMOINS', TRUE);
//Gain.SortGrid (0, FALSE);

TOB_CURetenu.Detail.Sort ('PKC_CUMULPAIE');
TOB_CURetenu.PutGridDetail (Retenu, FALSE, TRUE,
                            'PKC_CUMULPAIE;PKC_LIBELLE;PKC_PLUSMOINS', TRUE);
//Retenu.SortGrid (0, FALSE);

For i := 0 to TOB_CUGain_Herit.Detail.Count-1 do
    begin
    TOB_CumulsD:= TOB_CUGain_Herit.Detail[i];
    TOB_CumulsAutorisesD:= TOB_CumulsAutorises.FindFirst (['PEN_RUBRIQUE'],
                                                          [TOB_CumulsD.GetValue('PCR_CUMULPAIE')], True);

    if (TOB_CumulsAutorisesD<>nil) then
       TrueOrFalse:= 0
    else
       TrueOrFalse:= 1;

    if (TOB_CumulsD.FieldExists('AUTOR')) then
       begin
       TOB_CumulsD.PutValue ('AUTOR', TOB_CumulsAutorisesD<>nil);
       TOB_CumulsD.PutValue ('NUM', TrueOrFalse);
       end
    else
       begin
       TOB_CumulsD.AddChampSupValeur ('AUTOR', TOB_CumulsAutorisesD<>nil);
       TOB_CumulsD.AddChampSupValeur ('NUM', TrueOrFalse);
       end;
    end;
TOB_CUGain_Herit.Detail.Sort ('NUM;PCR_CUMULPAIE');
//TobDebug (TOB_CUGain_Herit);
TOB_CUGain_Herit.PutGridDetail (Gain_Herit, FALSE, TRUE,
                                'PCR_CUMULPAIE;PCR_LIBELLE;AUTOR', TRUE);
//Gain_Herit.SortGrid(0, False);
//Gain_Herit.SortGrid(2, True);

For i := 0 to TOB_CURetenu_Herit.Detail.Count-1 do
    begin
    TOB_CumulsD:= TOB_CURetenu_Herit.Detail[i];
    TOB_CumulsAutorisesD:= TOB_CumulsAutorises.FindFirst (['PEN_RUBRIQUE'],
                                                          [TOB_CumulsD.GetValue('PCR_CUMULPAIE')], True);

    if (TOB_CumulsAutorisesD<>nil) then
       TrueOrFalse:= 0
    else
       TrueOrFalse:= 1;

    if (TOB_CumulsD.FieldExists('AUTOR')) then
       begin
       TOB_CumulsD.PutValue ('AUTOR', TOB_CumulsAutorisesD<>nil);
       TOB_CumulsD.PutValue ('NUM', TrueOrFalse);
       end
    else
       begin
       TOB_CumulsD.AddChampSupValeur ('AUTOR', TOB_CumulsAutorisesD<>nil);
       TOB_CumulsD.AddChampSupValeur ('NUM', TrueOrFalse);
       end;
    end;
TOB_CURetenu_Herit.Detail.Sort ('NUM');
//TobDebug (TOB_CUGain_Herit);
TOB_CURetenu_Herit.PutGridDetail (Retenu_Herit, FALSE, TRUE,
                                  'PCR_CUMULPAIE;PCR_LIBELLE;AUTOR', TRUE);
//Retenu_Herit.SortGrid(0, False);
//Retenu_Herit.SortGrid(2, True);
}
For i := 0 to TOB_CURUB.Detail.Count-1 do
    begin
    TOBDetail:= TOB_CURUB.Detail[i];
    TOB_CumulsD:= TOB_Cumuls.FindFirst (['PCL_CUMULPAIE'],
                                        [TOBDetail.GetValue('PKC_CUMULPAIE')],
                                        True);
    if (TOBDetail.GetValue ('PKC_PLUSMOINS')='-') then
       Sens:='S'
    else
    if (TOBDetail.GetValue ('PKC_SENS')='+') then
       Sens:='P'
    else
    if (TOBDetail.GetValue ('PKC_SENS')='-') then
       Sens:='M';

{PT4
    TOB_CumulsD.AddChampSupValeur ('AUTOR', True);
}
    TOB_CumulsD.AddChampSupValeur ('AUTOR', 'X');
    TOB_CumulsD.AddChampSupValeur ('NUM', 0);
    TOB_CumulsD.AddChampSupValeur ('SENS', Sens);
    end;

For i := 0 to TOB_CURUB_Herit.Detail.Count-1 do
    begin
    TOBDetail:= TOB_CURUB_Herit.Detail[i];
    TOB_CumulsD:= TOB_Cumuls.FindFirst (['PCL_CUMULPAIE'],
                                        [TOBDetail.GetValue('PCR_CUMULPAIE')],
                                        True);
    if (TOB_CumulsD.FieldExists ('SENS')=False) then
       begin
       if (TOBDetail.GetValue ('PCR_SENS')='+') then
          Sens:='P'
       else
       if (TOBDetail.GetValue ('PCR_SENS')='-') then
          Sens:='M';
{PT4
       TOB_CumulsD.AddChampSupValeur ('AUTOR', False);
}
       TOB_CumulsD.AddChampSupValeur ('AUTOR', '-');
       TOB_CumulsD.AddChampSupValeur ('NUM', 1);
       Tob_CumulsD.AddChampSupValeur ('SENS', Sens);
       end;
    end;

For i := 0 to TOB_CumulsAutorises.Detail.Count-1 do
    begin
    TOBDetail:= TOB_CumulsAutorises.Detail[i];
    TOB_CumulsD:= TOB_Cumuls.FindFirst (['PCL_CUMULPAIE'],
                                        [TOBDetail.GetValue('PEN_RUBRIQUE')],
                                        True);
{PT4
    TOB_CumulsD.AddChampSupValeur ('AUTOR', True);
}
    TOB_CumulsD.AddChampSupValeur ('AUTOR', 'X');
    TOB_CumulsD.AddChampSupValeur ('NUM', 0);
    if (TOB_CumulsD.FieldExists ('SENS')=False) then
       begin
       Sens:='S';
       Tob_CumulsD.AddChampSupValeur ('SENS', Sens);
       end;
    end;

for i:= TOB_Cumuls.Detail.Count-1 downto 0 do
    begin
    TOB_CumulsD:= TOB_Cumuls.Detail [i];
    if (TOB_CumulsD.FieldExists ('SENS')=False) then
       FreeAndNil (TOB_CumulsD);
    end;

TOB_Cumuls.Detail.Sort ('NUM;PCL_CUMULPAIE');
TOB_Cumuls.PutGridDetail (Cumuls, FALSE, TRUE,
                          'PCL_CUMULPAIE;PCL_LIBELLE;AUTOR;SENS', TRUE);

end;
//FIN PT3

{PT3
procedure TOF_PGCUMULGESTIONSPEC.Test_Depose_Objet (Sender, Source: TObject;
                                                    X, Y: Integer;
                                                    State: TDragState;
                                                    var Accept: Boolean);
begin
Accept:= FALSE;
if (Sender is THgrid) then
    Accept:= TRUE;
end;


procedure TOF_PGCUMULGESTIONSPEC.Depose_Objet (Destination, Origine : TObject;
                                               X, Y: Integer);
var
T, TC, TD, TDD, P: TOB;
Prefixe, TypeP: string;
Nbre, i: Integer; // recherche du nombre d'objets selectionnes
begin
if (Destination = Origine) then
   Exit;

Nbre:= THGrid(Origine).RowCount-1;
if (Quoi = 'C') then
   TypeP:= 'CEG'
else
if (Quoi = 'S') then
   TypeP:= 'STD'
else
   TypeP:= 'DOS';

Modifie:= TRUE;
Retour:= TRUE;

for i:=1 to Nbre do
    begin
    if (THGrid(Origine).IsSelected(i)) then
       begin
// TD AU FORMAT CUMULPAIE
       TD:= TOB (THGrid(Origine).Objects[0, i]);

       if (TD=nil) then
          exit;
       if ((Origine<>Gain) and (Origine<>Retenu)) then
          begin
          if (Destination = Gain) then
             P:= TOB_CUGain
          else
          if (Destination = Retenu) then
             P:= TOB_CURetenu;

//Dans le cas où cumul non modifiable
//TobDebug (TOB_CumulsAutorises);
//TobDebug (TD);
          Prefixe:= Copy (TD.GetNomChamp(1), 1, 4);
          TDD:= TOB_CumulsAutorises.FindFirst (['PEN_RUBRIQUE'], [TD.GetValue (Prefixe+'CUMULPAIE')], True);

          if (TDD=nil) then
             begin
             PGIBox ('Vous n''êtes pas autorisé à modifier ce cumul.#13#10'+
                     'Seuls les cumuls personnalisés peuvent être modifiés.', //PT1
                     'Gestion spécifique');
             THGrid (Origine).ClearSelected;
             exit;
             end;

//Dans le cas où cumul déjà affecté
          TDD:= TOB_CUGain.FindFirst (['PKC_CUMULPAIE'],
                                      [TD.GetValue (Prefixe+'CUMULPAIE')], True);
          if (TDD<>nil) then
             begin
             PGIBox ('Cumul déjà affecté en gain', 'Gestion spécifique');
             THGrid (Origine).ClearSelected;
             exit;
             end;
          TDD:= TOB_CURetenu.FindFirst (['PKC_CUMULPAIE'],
                                        [TD.GetValue (Prefixe+'CUMULPAIE')], True);
          if (TDD<>nil) then
             begin
             PGIBox ('Cumul déjà affecté en retenue', 'Gestion spécifique');
             THGrid (Origine).ClearSelected;
             exit;
             end;

//Dans le cas où SOURCE = Cumuls
          if (Origine=Cumuls) then
             begin
             T:= TOB.Create ('CUMULRUBDOSSIER', P, -1);

             TD.ChangeParent (TOB_CumulCache, -1);
             TC:= TOB_Cumuls.FindFirst (['PCL_CUMULPAIE'], [TD.GetValue ('PCL_CUMULPAIE')], True);
             if (TC<>nil) then
                FreeAndNil (TC);

             T.PutValue ('PKC_CUMULPAIE', TD.GetValue ('PCL_CUMULPAIE'));
             T.PutValue ('PKC_LIBELLE', TD.GetValue ('PCL_LIBELLE'));
             T.PutValue ('PKC_RUBRIQUE', Rubrique);
             T.PutValue ('PKC_NATURERUB', Nature);
             if (Destination=Gain) then
                T.PutValue ('PKC_SENS', '+')
             else
                T.PutValue ('PKC_SENS', '-');
             T.PutValue ('PKC_PLUSMOINS', '+');
             end
          else
//Dans le cas où SOURCE = Gain_Herit
          if (Origine=Gain_Herit) then
             begin
             T:= TOB.Create ('CUMULRUBDOSSIER', P, -1);

             T.PutValue ('PKC_CUMULPAIE', TD.GetValue ('PCR_CUMULPAIE'));
             T.PutValue ('PKC_LIBELLE', TD.GetValue ('PCR_LIBELLE'));
             T.PutValue ('PKC_RUBRIQUE', Rubrique);
             T.PutValue ('PKC_NATURERUB', Nature);
             if (Destination=Gain) then
                begin
                T.PutValue ('PKC_SENS', '+');
                T.PutValue ('PKC_PLUSMOINS', '-');
                end
             else
                begin
                T.PutValue ('PKC_SENS', '-');
                T.PutValue ('PKC_PLUSMOINS', '+');
                end;
             end
          else
//Dans le cas où SOURCE = Retenu_Herit
          if (Origine=Retenu_Herit) then
             begin
             T:= TOB.Create ('CUMULRUBDOSSIER', P, -1);

             T.PutValue ('PKC_CUMULPAIE', TD.GetValue ('PCR_CUMULPAIE'));
             T.PutValue ('PKC_LIBELLE', TD.GetValue ('PCR_LIBELLE'));
             T.PutValue ('PKC_RUBRIQUE', Rubrique);
             T.PutValue ('PKC_NATURERUB', Nature);
             if (Destination=Gain) then
                begin
                T.PutValue ('PKC_SENS', '+');
                T.PutValue ('PKC_PLUSMOINS', '+');
                end
             else
                begin
                T.PutValue ('PKC_SENS', '-');
                T.PutValue ('PKC_PLUSMOINS', '-');
                end;
             end;
          end
       else
          begin
          if ((Destination=Cumuls) or (Destination=Gain_Herit) or
             (Destination=Retenu_Herit)) then
             begin
             if (Destination=Cumuls) then
                begin
                T:= TOB_CUGain_Herit.FindFirst (['PCR_CUMULPAIE'],
                                                [TD.GetValue('PKC_CUMULPAIE')], FALSE);
                if (T = nil) then
                   T:= TOB_CURetenu_Herit.FindFirst (['PCR_CUMULPAIE'],
                                                     [TD.GetValue('PKC_CUMULPAIE')], FALSE);

                if (T = nil) then
                   begin
                   T:= TOB.Create ('CUMULPAIE', TOB_Cumuls, -1);

                   T.PutValue ('PCL_CUMULPAIE', TD.GetValue ('PKC_CUMULPAIE'));
                   T.PutValue ('PCL_LIBELLE', TD.GetValue ('PKC_LIBELLE'));
                   end;

                FreeAndNil (TD);
                end;
             end
          else
             begin
             if (Destination = Gain) then
                P:= TOB_CUGain
             else
                P:= TOB_CURetenu;
             if (Origine=Gain) then
                begin
                if (Destination=Retenu) then
                   begin
                   TD.ChangeParent (P, -1);
                   TD.PutValue ('PKC_SENS', '-');
                   T:= TOB_CUGain_Herit.FindFirst (['PCR_CUMULPAIE'],
                                                   [TD.GetValue('PKC_CUMULPAIE')], FALSE);
                   if (T<>nil) then
                      TD.PutValue ('PKC_PLUSMOINS', '+')
                   else
                      begin
                      T:= TOB_CURetenu_Herit.FindFirst (['PCR_CUMULPAIE'],
                                                        [TD.GetValue('PKC_CUMULPAIE')], FALSE);
                      if (T<>nil) then
                         TD.PutValue ('PKC_PLUSMOINS', '-')
                      else
                         TD.PutValue ('PKC_PLUSMOINS', '+');
                      end;
                   end;
                end
             else
             if (Origine=Retenu) then
                begin
                if (Destination=Gain) then
                   begin
                   TD.ChangeParent (P, -1);
                   TD.PutValue ('PKC_SENS', '+');
                   T:= TOB_CUGain_Herit.FindFirst (['PCR_CUMULPAIE'],
                                                   [TD.GetValue('PKC_CUMULPAIE')], FALSE);
                   if (T<>nil) then
                      TD.PutValue ('PKC_PLUSMOINS', '-')
                   else
                      begin
                      T:= TOB_CURetenu_Herit.FindFirst (['PCR_CUMULPAIE'],
                                                        [TD.GetValue('PKC_CUMULPAIE')], FALSE);
                      if (T<>nil) then
                         TD.PutValue ('PKC_PLUSMOINS', '+')
                      else
                         TD.PutValue ('PKC_PLUSMOINS', '+');
                      end;
                   end;
                end;
             end;
          end;
       end;
    end;
THGrid (Origine).ClearSelected;
AfficheGrids;
end;


procedure TOF_PGCUMULGESTIONSPEC.CumulsMouseDown (Sender: TObject;
                                                  Button: TMouseButton;
                                                  Shift: TShiftState;
                                                  X, Y: Integer);
begin
if (Button = mbLeft) then
   begin
   if (Sender=Cumuls) then
      Cumuls.BeginDrag(TRUE, 10)
   else
   if (Sender=Gain) then
      Gain.BeginDrag(TRUE, 10)
   else
   if (Sender=Retenu) then
      Retenu.BeginDrag(TRUE, 10)
   else
   if (Sender=Gain_Herit) then
      Gain_Herit.BeginDrag(TRUE, 10)
   else
   if (Sender=Retenu_Herit) then
      Retenu_Herit.BeginDrag(TRUE, 10);
   end;
end;


procedure TOF_PGCUMULGESTIONSPEC.VireDejaUtilise;
var
T, TT : TOB;
i : Integer;
Sens : string;
begin
for i:= TOB_Cumuls.Detail.Count-1 downto 0 do
    begin
    T:= TOB_Cumuls.Detail[i];

    TT:= TOB_CURUB_Herit.FindFirst (['PCR_CUMULPAIE'],
                                    [T.GetValue('PCL_CUMULPAIE')], FALSE);
    if (TT<>nil) then
// ON AJOUTE DANS TOB_CumulCache
       T.ChangeParent (TOB_CumulCache, 0);

    TT:= TOB_CURUB.FindFirst (['PKC_CUMULPAIE'],
                              [T.GetValue('PCL_CUMULPAIE')], FALSE);
    if (TT<>nil) then
// ON AJOUTE DANS TOB_CumulCache
       T.ChangeParent (TOB_CumulCache, 0);
    end;

for i:= TOB_CURUB_Herit.Detail.Count-1 downto 0 do
    begin
    T:= TOB_CURUB_Herit.Detail[I];

    SENS:= T.GetValue ('PCR_SENS');

    TT:= TOB_CumulCache.FindFirst (['PCL_CUMULPAIE'],
                                   [T.GetValue('PCR_CUMULPAIE')], FALSE);
    if (TT<>nil) then
       begin
       if (SENS='+') then
          T.Changeparent (TOB_CUGain_Herit, 0)
       else
          T.Changeparent (TOB_CURetenu_Herit, 0);
       end;
    end;
FreeAndNil (TOB_CURUB_Herit);

for i:= TOB_CURUB.Detail.Count-1 downto 0 do
    begin
    T:= TOB_CURUB.Detail[I];

    SENS:= T.GetValue ('PKC_SENS');

    if (SENS='+') then
       T.Changeparent (TOB_CUGain, 0)
    else
       T.Changeparent (TOB_CURetenu, 0);
    end;
FreeAndNil (TOB_CURUB);
end;
}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 22/02/2007
Modifié le ... :   /  /
Description .. : Enregistrement des modifications
Mots clefs ... : TOF;PGCUMULGESTIONSPEC;PAIE;GESTIONSPECIFIQUE
*****************************************************************}
procedure TOF_PGCUMULGESTIONSPEC.Sauve;
begin
if (not Modifie) then
   exit;
ExecuteSQL ('DELETE FROM CUMULRUBDOSSIER WHERE'+
            ' PKC_NATURERUB="'+Nature+'" AND'+
            ' PKC_RUBRIQUE="'+Rubrique+'"');
{PT3
TOB_CUGain.SetAllModifie (TRUE);
TOB_CURetenu.SetAllModifie (TRUE);
TOB_CUGain.InsertDB (nil, TRUE);
TOB_CURetenu.InsertDB (nil, TRUE);
}
TOB_CURUB.SetAllModifie (TRUE);
TOB_CURUB.InsertDB (nil, TRUE);
//FIN PT3
Modifie:= FALSE;
end;

{PT3
procedure TOF_PGCUMULGESTIONSPEC.BNonAffecte_GainClick(Sender: TObject);
begin
Depose_Objet (Gain, Cumuls, 0, 0);
Depose_Objet (Gain, Gain_Herit, 0, 0);
Depose_Objet (Gain, Retenu_Herit, 0, 0);
end;

procedure TOF_PGCUMULGESTIONSPEC.BNonAffecte_RetenuClick(Sender: TObject);
begin
Depose_Objet (Retenu, Cumuls, 0, 0);
Depose_Objet (Retenu, Gain_Herit, 0, 0);
Depose_Objet (Retenu, Retenu_Herit, 0, 0);
end;

procedure TOF_PGCUMULGESTIONSPEC.BGain_NonAffecteClick(Sender: TObject);
begin
Depose_Objet (Cumuls, Gain, 0, 0);
Depose_Objet (Gain_Herit, Gain, 0, 0);
Depose_Objet (Retenu_Herit, Gain, 0, 0);
end;

procedure TOF_PGCUMULGESTIONSPEC.BRetenu_NonAffecteClick(Sender: TObject);
begin
Depose_Objet (Cumuls, Retenu, 0, 0);
Depose_Objet (Gain_Herit, Retenu, 0, 0);
Depose_Objet (Retenu_Herit, Retenu, 0, 0);
end;

procedure TOF_PGCUMULGESTIONSPEC.BGain_RetenuClick(Sender: TObject);
begin
Depose_Objet (Retenu, Gain, 0, 0);
end;

procedure TOF_PGCUMULGESTIONSPEC.BRetenu_GainClick(Sender: TObject);
begin
Depose_Objet (Gain, Retenu, 0, 0);
end;
}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 22/02/2007
Modifié le ... :   /  /
Description .. : Double-clic sur la grille des cumuls non affectés
Mots clefs ... : TOF;PGCUMULGESTIONSPEC;PAIE;GESTIONSPECIFIQUE
*****************************************************************}
procedure TOF_PGCUMULGESTIONSPEC.CumulsDblClick(Sender: TObject);
begin
AglLanceFiche ('PAY', 'CUMUL', '', '', 'ACTION=MODIFICATION');
end;

//PT3
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 16/07/2007
Modifié le ... :   /  /
Description .. : Sortie d'une cellule de la grille des cumuls
Mots clefs ... : TOF;PGCUMULGESTIONSPEC;PAIE;GESTIONSPECIFIQUE
*****************************************************************}
procedure TOF_PGCUMULGESTIONSPEC.CumulsCellExit (Sender: TObject;
                                                 var ACol,ARow: Integer;
                                                 var Cancel: Boolean);
var
TDetailHerit, TDetailPerso : TOB;
Sens : String;
begin
if (ACol=3) then
   begin
   if (Cumuls.CellValues[2, ARow]= 'X') then
      begin
      TDetailHerit:= TOB_CURUB_Herit.FindFirst (['PCR_CUMULPAIE'],
                                                [Cumuls.CellValues[0, ARow]],
                                                True);
      TDetailPerso:= TOB_CURUB.FindFirst (['PKC_CUMULPAIE'],
                                          [Cumuls.CellValues[0, ARow]], True);
      Sens:= Cumuls.CellValues[ACol, ARow];
//Il existe un héritage
      if (TDetailHerit<>nil) then
         begin
//Héritage en gain
         if (TDetailHerit.GetValue ('PCR_SENS')='+') then
            begin
            if ((Sens='P') and (TDetailPerso<>nil)) then
               begin
               FreeAndNil (TDetailPerso);
               Modifie:= True;
               end
            else
            if (Sens='S') then
               begin
               if (TDetailPerso=nil) then
                  begin
                  TDetailPerso:= TOB.Create ('CUMULRUBDOSSIER', TOB_CURUB, -1);
                  TDetailPerso.AddChampSupValeur ('PKC_NATURERUB', Nature);
                  TDetailPerso.AddChampSupValeur ('PKC_RUBRIQUE', Rubrique);
                  TDetailPerso.AddChampSupValeur ('PKC_CUMULPAIE',
                                                  Cumuls.CellValues[0, ARow]);
                  TDetailPerso.AddChampSupValeur ('PKC_ETABLISSEMENT', '');
                  TDetailPerso.AddChampSupValeur ('PKC_LIBELLE',
                                                  Cumuls.CellValues[1, ARow]);
                  end;
               TDetailPerso.PutValue ('PKC_SENS', '+');
               TDetailPerso.PutValue ('PKC_PLUSMOINS', '-');
               Modifie:= True;
               end
            else
            if (Sens='M') then
               begin
               if (TDetailPerso=nil) then
                  begin
                  TDetailPerso:= TOB.Create ('CUMULRUBDOSSIER', TOB_CURUB, -1);
                  TDetailPerso.AddChampSupValeur ('PKC_NATURERUB', Nature);
                  TDetailPerso.AddChampSupValeur ('PKC_RUBRIQUE', Rubrique);
                  TDetailPerso.AddChampSupValeur ('PKC_CUMULPAIE',
                                                  Cumuls.CellValues[0, ARow]);
                  TDetailPerso.AddChampSupValeur ('PKC_ETABLISSEMENT', '');
                  TDetailPerso.AddChampSupValeur ('PKC_LIBELLE',
                                                  Cumuls.CellValues[1, ARow]);
                  end;
               TDetailPerso.PutValue ('PKC_SENS', '-');
               TDetailPerso.PutValue ('PKC_PLUSMOINS', '+');
               Modifie:= True;
               end;
            end
         else
//Héritage en retenue
            begin
            if ((Sens='M') and (TDetailPerso<>nil)) then
               begin
               FreeAndNil (TDetailPerso);
               Modifie:= True;
               end
            else
            if (Sens='S') then
               begin
               if (TDetailPerso=nil) then
                  begin
                  TDetailPerso:= TOB.Create ('CUMULRUBDOSSIER', TOB_CURUB, -1);
                  TDetailPerso.AddChampSupValeur ('PKC_NATURERUB', Nature);
                  TDetailPerso.AddChampSupValeur ('PKC_RUBRIQUE', Rubrique);
                  TDetailPerso.AddChampSupValeur ('PKC_CUMULPAIE',
                                                  Cumuls.CellValues[0, ARow]);
                  TDetailPerso.AddChampSupValeur ('PKC_ETABLISSEMENT', '');
                  TDetailPerso.AddChampSupValeur ('PKC_LIBELLE',
                                                  Cumuls.CellValues[1, ARow]);
                  end;
               TDetailPerso.PutValue ('PKC_SENS', '-');
               TDetailPerso.PutValue ('PKC_PLUSMOINS', '-');
               Modifie:= True;
               end
            else
            if (Sens='P') then
               begin
               if (TDetailPerso=nil) then
                  begin
                  TDetailPerso:= TOB.Create ('CUMULRUBDOSSIER', TOB_CURUB, -1);
                  TDetailPerso.AddChampSupValeur ('PKC_NATURERUB', Nature);
                  TDetailPerso.AddChampSupValeur ('PKC_RUBRIQUE', Rubrique);
                  TDetailPerso.AddChampSupValeur ('PKC_CUMULPAIE',
                                                  Cumuls.CellValues[0, ARow]);
                  TDetailPerso.AddChampSupValeur ('PKC_ETABLISSEMENT', '');
                  TDetailPerso.AddChampSupValeur ('PKC_LIBELLE',
                                                  Cumuls.CellValues[1, ARow]);
                  end;
               TDetailPerso.PutValue ('PKC_SENS', '+');
               TDetailPerso.PutValue ('PKC_PLUSMOINS', '+');
               Modifie:= True;
               end;
            end;
         end
      else
//Il n'existe pas d'héritage
         begin
         if ((Sens='S') and (TDetailPerso<>nil)) then
            begin
            FreeAndNil (TDetailPerso);
            Modifie:= True;
            end
         else
         if (Sens='P') then
            begin
            if (TDetailPerso=nil) then
               begin
               TDetailPerso:= TOB.Create ('CUMULRUBDOSSIER', TOB_CURUB, -1);
               TDetailPerso.AddChampSupValeur ('PKC_NATURERUB', Nature);
               TDetailPerso.AddChampSupValeur ('PKC_RUBRIQUE', Rubrique);
               TDetailPerso.AddChampSupValeur ('PKC_CUMULPAIE',
                                               Cumuls.CellValues[0, ARow]);
               TDetailPerso.AddChampSupValeur ('PKC_ETABLISSEMENT', '');
               TDetailPerso.AddChampSupValeur ('PKC_LIBELLE',
                                               Cumuls.CellValues[1, ARow]);
               end;
            TDetailPerso.PutValue ('PKC_SENS', '+');
            TDetailPerso.PutValue ('PKC_PLUSMOINS', '+');
            Modifie:= True;
            end
         else
         if (Sens='M') then
            begin
            if (TDetailPerso=nil) then
               begin
               TDetailPerso:= TOB.Create ('CUMULRUBDOSSIER', TOB_CURUB, -1);
               TDetailPerso.AddChampSupValeur ('PKC_NATURERUB', Nature);
               TDetailPerso.AddChampSupValeur ('PKC_RUBRIQUE', Rubrique);
               TDetailPerso.AddChampSupValeur ('PKC_CUMULPAIE',
                                               Cumuls.CellValues[0, ARow]);
               TDetailPerso.AddChampSupValeur ('PKC_ETABLISSEMENT', '');
               TDetailPerso.AddChampSupValeur ('PKC_LIBELLE',
                                               Cumuls.CellValues[1, ARow]);
               end;
            TDetailPerso.PutValue ('PKC_SENS', '-');
            TDetailPerso.PutValue ('PKC_PLUSMOINS', '+');
            Modifie:= True;
            end;
         end;
      end;
   end;
end;
//FIN PT3

{PT3
procedure TOF_PGCUMULGESTIONSPEC.GainDblClick(Sender: TObject);
begin
AglLanceFiche ('PAY', 'CUMUL', '', Gain.Cells[0, Gain.Row],
               'ACTION=MODIFICATION');
end;


procedure TOF_PGCUMULGESTIONSPEC.RetenuDblClick(Sender: TObject);
begin
AglLanceFiche ('PAY', 'CUMUL', '', Retenu.Cells[0, Retenu.Row],
               'ACTION=MODIFICATION');
end;

}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 08/03/2007
Modifié le ... :   /  /
Description .. : Click sur le bouton Imprimer
Mots clefs ... : TOF;PGCUMULGESTIONSPEC;PAIE;GESTIONSPECIFIQUE
*****************************************************************}
procedure TOF_PGCUMULGESTIONSPEC.Impression(Sender: TObject);
var
TCR, TCRD, TCRDF, TCRF, Tob_Cumul, Tob_CumulD, TobEtat, TobEtatD : TOB;
St, StPlus, StPluss : string;
Q : TQuery;
i, MaxCum : integer;
begin
Sauve;    //PT2-3
StPlus:= 'WHERE ##PCR_PREDEFINI## PCR_NATURERUB="'+GetControlText ('LNATURERUB')+'" AND'+
         ' PCR_RUBRIQUE="'+GetControlText ('LRUBRIQUE')+'"';

{Optimisation - PT2-2
Q:= OpenSQL ('SELECT DISTINCT PCR_NATURERUB, PCR_RUBRIQUE, PRM_ABREGE,'+
             ' PCT_ABREGE, PCT_PREDEFINI, PRM_PREDEFINI, CO_LIBELLE'+
             ' FROM CUMULRUBRIQUE'+
             ' LEFT JOIN COTISATION ON'+
             ' PCR_NATURERUB=PCT_NATURERUB AND'+
             ' ##PCT_PREDEFINI## PCR_RUBRIQUE=PCT_RUBRIQUE'+
             ' LEFT JOIN REMUNERATION ON'+
             ' PCR_NATURERUB=PRM_NATURERUB AND'+
             ' ##PRM_PREDEFINI## PCR_RUBRIQUE=PRM_RUBRIQUE'+
             ' LEFT JOIN COMMUN ON'+
             ' CO_TYPE="NTR" AND CO_CODE=PCR_NATURERUB '+StPlus, True);
TobEtat:= Tob.Create('La Rubriques', nil, -1);
TobEtat.LoadDetailDB('La Rubriques', '', '', Q, False);
Ferme(Q);
}
St:= 'SELECT DISTINCT PCR_NATURERUB, PCR_RUBRIQUE, PRM_ABREGE, PCT_ABREGE,'+
     ' PCT_PREDEFINI, PRM_PREDEFINI, CO_LIBELLE'+
     ' FROM CUMULRUBRIQUE'+
     ' LEFT JOIN COTISATION ON'+
     ' PCR_NATURERUB=PCT_NATURERUB AND'+
     ' ##PCT_PREDEFINI## PCR_RUBRIQUE=PCT_RUBRIQUE'+
     ' LEFT JOIN REMUNERATION ON'+
     ' PCR_NATURERUB=PRM_NATURERUB AND'+
     ' ##PRM_PREDEFINI## PCR_RUBRIQUE=PRM_RUBRIQUE'+
     ' LEFT JOIN COMMUN ON'+
     ' CO_TYPE="NTR" AND CO_CODE=PCR_NATURERUB '+StPlus;
TobEtat:= Tob.Create ('La Rubriques', nil, -1);
TobEtat.LoadDetailDBFromSQL ('La Rubriques', St);
//FIN PT2-2

StPluss:= 'WHERE PKC_NATURERUB="'+GetControlText ('LNATURERUB')+'" AND'+
         ' PKC_RUBRIQUE="'+GetControlText ('LRUBRIQUE')+'"';

{Optimisation - PT2-2
Q:= OpenSQL ('SELECT DISTINCT PKC_NATURERUB AS PCR_NATURERUB,'+
             ' PKC_RUBRIQUE AS PCR_RUBRIQUE, PRM_ABREGE, PCT_ABREGE,'+
             ' PCT_PREDEFINI, PRM_PREDEFINI, CO_LIBELLE'+
             ' FROM CUMULRUBDOSSIER'+
             ' LEFT JOIN COTISATION ON'+
             ' PKC_NATURERUB=PCT_NATURERUB AND'+
             ' ##PCT_PREDEFINI## PKC_RUBRIQUE=PCT_RUBRIQUE'+
             ' LEFT JOIN REMUNERATION ON'+
             ' PKC_NATURERUB=PRM_NATURERUB AND'+
             ' ##PRM_PREDEFINI## PKC_RUBRIQUE=PRM_RUBRIQUE'+
             ' LEFT JOIN COMMUN ON'+
             ' CO_TYPE="NTR" AND CO_CODE=PKC_NATURERUB '+StPluss, True);
TCRD:= Tob.Create('La Rubriques', nil, -1);
TCRD.LoadDetailDB('La Rubriques', '', '', Q, False);
Ferme(Q);
}
St:= 'SELECT DISTINCT PKC_NATURERUB AS PCR_NATURERUB,'+
     ' PKC_RUBRIQUE AS PCR_RUBRIQUE, PRM_ABREGE, PCT_ABREGE, PCT_PREDEFINI,'+
     ' PRM_PREDEFINI, CO_LIBELLE'+
     ' FROM CUMULRUBDOSSIER'+
     ' LEFT JOIN COTISATION ON'+
     ' PKC_NATURERUB=PCT_NATURERUB AND'+
     ' ##PCT_PREDEFINI## PKC_RUBRIQUE=PCT_RUBRIQUE'+
     ' LEFT JOIN REMUNERATION ON'+
     ' PKC_NATURERUB=PRM_NATURERUB AND'+
     ' ##PRM_PREDEFINI## PKC_RUBRIQUE=PRM_RUBRIQUE'+
     ' LEFT JOIN COMMUN ON'+
     ' CO_TYPE="NTR" AND CO_CODE=PKC_NATURERUB '+StPluss;
TCRD:= Tob.Create ('La Rubriques', nil, -1);
TCRD.LoadDetailDBFromSQL ('La Rubriques', St);
//FIN PT2-2

TCRDF:= TCRD.FindFirst ([''], [''], False);
{PT2-1
While (TCRDD<>nil) do
      begin
      TobEtatD:= TobEtat.FindFirst (['PCR_NATURERUB', 'PCR_RUBRIQUE'],
                                    [TCRDD.GetValue ('PCR_NATURERUB'),
                                     TCRDD.GetValue ('PCR_RUBRIQUE')], False);
      if (Assigned (TobEtatD)) then
         FreeAndNil(TCRDD)
      else
         TCRDD.ChangeParent (TobEtat, -1);

      TCRDD:= TCRD.FindFirst ([''], [''], False);
      end;
}
While (TCRDF<>nil) do
      begin
      TobEtatD:= TobEtat.FindFirst (['PCR_NATURERUB', 'PCR_RUBRIQUE'],
                                    [TCRDF.GetValue ('PCR_NATURERUB'),
                                     TCRDF.GetValue ('PCR_RUBRIQUE')], False);
      if (Assigned (TobEtatD)) then
         FreeAndNil(TCRDF)
      else
         TCRDF.ChangeParent (TobEtat, -1);

      TCRDF:= TCRD.FindFirst ([''], [''], False);
      end;
//FIN PT2-1
FreeAndNil (TCRD);

{PT2-1
if (TobEtat.Detail.Count=0) then
   FreeAndNil (TobEtat);
}

Q:= OpenSql ('SELECT COUNT (*) NBRE'+
             ' FROM CUMULPAIE WHERE'+
             ' PCL_CUMULPAIE <> "" AND ##PCL_PREDEFINI##', True);
if (not Q.EOF) then
   MaxCum:= Q.FindField ('NBRE').asInteger
else
   exit;
Ferme(Q);

TobEtatD:= TobEtat.FindFirst ([''], [''], False);
if (TobEtatD<>nil) then      //PT2-1
   begin
   TobEtatD.AddChampSupValeur ('TYPEINFO', 'LIBELLE', True);
   for i:= 1 to MaxCum do
       TobEtatD.AddChampSup ('C'+IntToStr (i), True);
   end;

Tob_Cumul:= Tob.Create('Les cumuls', nil, -1);
Tob_cumul.LoadDetailDB ('CUMULPAIE', '', '', nil, False);

{Optimisation - PT2-2
Q:= OpenSQL ('SELECT DISTINCT PCR_NATURERUB, PCR_RUBRIQUE, PCR_CUMULPAIE,'+
             ' PCR_LIBELLE, PCR_SENS'+
             ' FROM CUMULRUBRIQUE '+StPlus, True);
TCR:= Tob.Create('La Tob chargee', nil, -1);
TCR.LoadDetailDB('Mon CUMULRUBRIQUE', '', '', Q, False);
Ferme(Q);
}
St:= 'SELECT DISTINCT PCR_NATURERUB, PCR_RUBRIQUE, PCR_CUMULPAIE, PCR_LIBELLE,'+
     ' PCR_SENS'+
     ' FROM CUMULRUBRIQUE '+StPlus;
TCR:= Tob.Create ('La Tob chargee', nil, -1);
TCR.LoadDetailDBFromSQL ('Mon CUMULRUBRIQUE', St);
//FIN PT2-2

{Optimisation - PT2-2
Q:= OpenSQL ('SELECT DISTINCT PKC_NATURERUB AS PCR_NATURERUB,'+
             ' PKC_RUBRIQUE AS PCR_RUBRIQUE, PKC_CUMULPAIE AS PCR_CUMULPAIE,'+
             ' PKC_LIBELLE AS PCR_LIBELLE, PKC_SENS AS PCR_SENS'+
             ' FROM CUMULRUBDOSSIER '+StPluss, True);
TCRD:= Tob.Create('La Tob chargee', nil, -1);
TCRD.LoadDetailDB('Mon CUMULRUBRIQUE', '', '', Q, False);
Ferme(Q);
}
St:= 'SELECT DISTINCT PKC_NATURERUB AS PCR_NATURERUB,'+
     ' PKC_RUBRIQUE AS PCR_RUBRIQUE, PKC_CUMULPAIE AS PCR_CUMULPAIE,'+
     ' PKC_LIBELLE AS PCR_LIBELLE, PKC_SENS AS PCR_SENS'+
     ' FROM CUMULRUBDOSSIER '+StPluss;
TCRD:= Tob.Create('La Tob chargee', nil, -1);
TCRD.LoadDetailDBFromSQL('Mon CUMULRUBRIQUE', St);
//FIN PT2-2

if (GetControlText ('LNATURERUB')='AAA') then
   begin
   TobEtatD:= Tob.Create ('Les Rubriques', TobEtat, -1);
   TobEtatD.AddChampSupValeur ('PCR_NATURERUB', 'AAA');
   TobEtatD.AddChampSupValeur ('PCR_RUBRIQUE', '0000');
   TobEtatD.AddChampSupValeur ('PRM_ABREGE', '');
   TobEtatD.AddChampSupValeur ('PCT_ABREGE', '');
   TobEtatD.AddChampSupValeur ('PCT_PREDEFINI', '');
   TobEtatD.AddChampSupValeur ('PRM_PREDEFINI', '');
   TobEtatD.AddChampSupValeur ('CO_LIBELLE', 'Rémunération');
   TobEtatD.AddChampSupValeur ('TYPEINFO', 'Rémunération');
   for i:= 1 to MaxCum do
       TobEtatD.AddChampSup ('C'+IntToStr (i), False);
   end
else
if (GetControlText ('LNATURERUB')='BAS') then
   begin
   TobEtatD:= Tob.Create ('Les Rubriques', TobEtat, -1);
   TobEtatD.AddChampSupValeur ('PCR_NATURERUB', 'BAS');
   TobEtatD.AddChampSupValeur ('PCR_RUBRIQUE', '0000');
   TobEtatD.AddChampSupValeur ('PRM_ABREGE', '');
   TobEtatD.AddChampSupValeur ('PCT_ABREGE', '');
   TobEtatD.AddChampSupValeur ('PCT_PREDEFINI', '');
   TobEtatD.AddChampSupValeur ('PRM_PREDEFINI', '');
   TobEtatD.AddChampSupValeur ('CO_LIBELLE', 'Base de cotisation');
   TobEtatD.AddChampSupValeur ('TYPEINFO', 'Base de cotisation');
   for i:= 1 to MaxCum do
       TobEtatD.AddChampSup ('C'+IntToStr (i), False);
   end
else
if (GetControlText ('LNATURERUB')='COT') then
   begin
   TobEtatD:= Tob.Create ('Les Rubriques', TobEtat, -1);
   TobEtatD.AddChampSupValeur ('PCR_NATURERUB', 'COT');
   TobEtatD.AddChampSupValeur ('PCR_RUBRIQUE', '0000');
   TobEtatD.AddChampSupValeur ('PRM_ABREGE', '');
   TobEtatD.AddChampSupValeur ('PCT_ABREGE', '');
   TobEtatD.AddChampSupValeur ('PCT_PREDEFINI', '');
   TobEtatD.AddChampSupValeur ('PRM_PREDEFINI', '');
   TobEtatD.AddChampSupValeur ('CO_LIBELLE', 'Cotisation');
   TobEtatD.AddChampSupValeur ('TYPEINFO', 'Cotisation');
   for i:= 1 to MaxCum do
       TobEtatD.AddChampSup ('C'+IntToStr (i), False);
   end;

TobEtat.Detail.Sort('PCR_NATURERUB;PCR_RUBRIQUE;TYPEINFO');

TCRDF:= TCRD.FindFirst ([''], [''], False);
While (TCRDF<>nil) do
      begin
      TCRF:= TCR.FindFirst (['PCR_NATURERUB', 'PCR_RUBRIQUE', 'PCR_CUMULPAIE'],
                            [TCRDF.GetValue ('PCR_NATURERUB'),
                             TCRDF.GetValue ('PCR_RUBRIQUE'),
                             TCRDF.GetValue ('PCR_CUMULPAIE')], False);
      if (Assigned (TCRF)) then
         begin
         if (TCRF.GetValue ('PCR_SENS')<>TCRDF.GetValue('PCR_SENS')) then
            begin
            FreeAndNil(TCRF);
            TCRDF.ChangeParent (TCR, -1);
            end
         else
            begin
            FreeAndNil(TCRF);
            FreeAndNil(TCRDF);
            end;
         end
      else
         TCRDF.ChangeParent (TCR, -1);

      TCRDF:= TCRD.FindFirst ([''], [''], False);
      end;

TCR.Detail.Sort('PCR_NATURERUB;PCR_RUBRIQUE;PCR_CUMULPAIE');

TobEtatD:= TobEtat.FindFirst ([''], [''], False);
While (TobEtatD<>nil) do
      begin
      if (TobEtatD.GetValue ('PCR_RUBRIQUE')='0000') then
         begin
         for i:= 1 to Tob_Cumul.Detail.Count do
             TobEtatD.PutValue ('C'+IntToStr (i),
                                Tob_Cumul.Detail[i-1].GetValue ('PCL_CUMULPAIE'));
         end;
      if (TCR <> nil) then
         TCRF:= TCR.FindFirst (['PCR_NATURERUB', 'PCR_RUBRIQUE'],
                               [TobEtatD.GetValue ('PCR_NATURERUB'),
                                TobEtatD.GetValue ('PCR_RUBRIQUE')], FALSE);
      if (TCRF <> nil) then
         begin
         while TCRF <> nil do
               begin
               Tob_CumulD:= Tob_Cumul.FindFirst (['PCL_CUMULPAIE'],
                                                 [TCRF.GetValue ('PCR_CUMULPAIE')], False);
               i:= Tob_CumulD.GetIndex;
               TobEtatD.PutValue ('C'+IntToStr (i+1), TCRF.GetValue ('PCR_SENS'));
               TCRF:= TCR.FindNext (['PCR_NATURERUB', 'PCR_RUBRIQUE'],
                                    [TobEtatD.GetValue ('PCR_NATURERUB'),
                                     TobEtatD.GetValue ('PCR_RUBRIQUE')], FALSE);
               end;
         end;
      TobEtatD:= TobEtat.FindNext ([''], [''], False);
      end;

FreeAndNil (Tob_Cumul);
FreeAndNil (TCR);
FreeAndNil (TCRD);

LanceEtatTob ('E', 'PGA', 'PGR', TobEtat, True, False, False,
              TPageControl(GetControl('Pages')), '', Ecran.Caption, False);
FreeAndNil (TobEtat);
end;

Initialization
  registerclasses ([TOF_PGCUMULGESTIONSPEC]);
end.
