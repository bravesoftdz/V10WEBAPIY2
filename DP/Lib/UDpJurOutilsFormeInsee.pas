{***********UNITE*************************************************
Auteur  ...... : BM
Créé le ...... : 16/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
Unit UDpJurOutilsFormeInsee;
//////////////////////////////////////////////////////////////////
Interface
//////////////////////////////////////////////////////////////////
Uses
   comctrls, UTOB, ed_tools, Forms, UFImgListe;

//////////////////////////////////////////////////////////////////
Type
   TForme = record
      sCodeInsee_c  : string;
      iNiveau_c     : integer;
      sLibelle_c    : string;
      sNonActif_c   : string;
      sForme_c      : string;
      sFormePrive_c : string;
      sFormeSte_c   : string;
      sFormeSci_c   : string;
      sFormeAsso_c  : string;
      sCoop_c       : string;
      sRegleFisc_c  : string;
      sSectionBnc_c : string;
end;

Type   TPForme = ^TForme;


function  TWLongueurCodeInsee(iNiveau_p : integer) : integer;
function  TWNiveauCodeInsee(sCodeInsee_p : string) : integer;
procedure TWInitListeIcones(tvArbre_p : TTreeView);
procedure TWChargeFormes(tvArbre_p : TTreeView);
function  TWFormateLib(sForme_p, sCodeInsee_p, sLibelle_p : string) : string;
function  TWIcone(sNonActif_p : string; iNiveau_p : integer) : integer;
procedure TWVideFormes(tvArbre_p : TTreeView);

procedure OnExpand(tvArbre_p : TTreeView);
procedure OnCollapse(tvArbre_p : TTreeView);

function  GetCodeInseeFromForme(sForme_p : string;
                                var sCodeInsee_p, sFormePrive_p, sFormeSte_p,
                                sFormeSci_p, sCoop_p : string) : boolean;

//////////////////////////////////////////////////////////////////
Implementation
//////////////////////////////////////////////////////////////////

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 15/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TWLongueurCodeInsee(iNiveau_p : integer) : integer;
begin
   if iNiveau_p = 3 then
      result := 4
   else
      result := iNiveau_p;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 28/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function  TWNiveauCodeInsee(sCodeInsee_p : string) : integer;
var
   iLg_l : integer;
begin
   iLg_l := Length(sCodeInsee_p);
   if iLg_l = 4 then
      result := 3
   else
      result := iLg_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 28/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TWInitListeIcones(tvArbre_p : TTreeView);
begin
   if FImgListe = nil then
      Application.CreateForm(TFImgListe, FImgListe);

   tvArbre_p.Images := FImgListe.NiveauxInsee;
   tvArbre_p.StateImages := FImgListe.NiveauxInsee;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 15/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TWChargeFormes(tvArbre_p : TTreeView);
var
   OBFormeInsee_l : TOB;
   LaForme_l : TPForme;
   iNbForme_l, iNoeud_l : integer;
   aotnNoeux_l : array [0..2] of TTreeNode;
begin
   TWVideFormes(tvArbre_p);

   tvArbre_p.Items.BeginUpdate;
   OBFormeInsee_l := TOB.Create('YFORMESINSEE', nil, -1);
   OBFormeInsee_l.LoadDetailDBFromSQL('YFORMESINSEE',
                                      'SELECT * FROM YFORMESINSEE ' +
                                      'WHERE YFJ_CODEINSEE NOT LIKE "58%" ' +
                                      'ORDER BY YFJ_CODEINSEE');

   aotnNoeux_l[0] := nil;
   aotnNoeux_l[1] := nil;
   aotnNoeux_l[2] := nil;
   iNoeud_l := 1;

   New(LaForme_l);
   LaForme_l^.sCodeInsee_c  := '';
   LaForme_l^.sForme_c      := '';
   LaForme_l^.sFormePrive_c := '';
   LaForme_l^.sFormeSte_c   := '';
   LaForme_l^.sFormeSci_c   := '';
   LaForme_l^.sFormeAsso_c  := '';
   LaForme_l^.iNiveau_c     := 1;
   LaForme_l^.sNonActif_c   := '-';
   LaForme_l^.sCoop_c       := '-';
   LaForme_l^.sRegleFisc_c  := '';
   LaForme_l^.sSectionBnc_c := '';
   LaForme_l^.sLibelle_c := 'Aucune';

   aotnNoeux_l[LaForme_l.iNiveau_c - 1] := tvArbre_p.Items.AddObject(aotnNoeux_l[LaForme_l.iNiveau_c - 1],
                                                LaForme_l^.sLibelle_c, LaForme_l);

   aotnNoeux_l[LaForme_l.iNiveau_c - 1].ImageIndex := TWIcone(LaForme_l^.sNonActif_c, LaForme_l^.iNiveau_c);
   aotnNoeux_l[LaForme_l.iNiveau_c - 1].SelectedIndex := TWIcone(LaForme_l^.sNonActif_c, LaForme_l^.iNiveau_c);

   iNbForme_l := 0;
   while iNbForme_l < OBFormeInsee_l.Detail.Count do
   begin
      New(LaForme_l);
      LaForme_l^.sCodeInsee_c  := OBFormeInsee_l.Detail[iNbForme_l].GetString('YFJ_CODEINSEE');
      LaForme_l^.sForme_c      := OBFormeInsee_l.Detail[iNbForme_l].GetString('YFJ_FORME');
      LaForme_l^.sFormePrive_c := OBFormeInsee_l.Detail[iNbForme_l].GetString('YFJ_FORMEGRPPRIVE');
      LaForme_l^.sFormeSte_c   := OBFormeInsee_l.Detail[iNbForme_l].GetString('YFJ_FORMESTE');
      LaForme_l^.sFormeSci_c   := OBFormeInsee_l.Detail[iNbForme_l].GetString('YFJ_FORMESCI');
      LaForme_l^.sFormeAsso_c  := OBFormeInsee_l.Detail[iNbForme_l].GetString('YFJ_FORMEASSO');
      LaForme_l^.iNiveau_c     := OBFormeInsee_l.Detail[iNbForme_l].GetInteger('YFJ_NIVEAU');
      LaForme_l^.sNonActif_c   := OBFormeInsee_l.Detail[iNbForme_l].GetString('YFJ_NONACTIF');
      LaForme_l^.sCoop_c       := OBFormeInsee_l.Detail[iNbForme_l].GetString('YFJ_COOP');
      LaForme_l^.sRegleFisc_c  := OBFormeInsee_l.Detail[iNbForme_l].GetString('YFJ_REGLEFISC');
      LaForme_l^.sSectionBnc_c := OBFormeInsee_l.Detail[iNbForme_l].GetString('YFJ_SECTIONBNC');

      LaForme_l^.sLibelle_c := TWFormateLib(LaForme_l^.sForme_c,
                                           LaForme_l^.sCodeInsee_c,
                                           OBFormeInsee_l.Detail[iNbForme_l].GetString('YFJ_LIBELLE'));

      if LaForme_l.iNiveau_c <= iNoeud_l then
      begin
         aotnNoeux_l[LaForme_l.iNiveau_c - 1] := tvArbre_p.Items.AddObject(aotnNoeux_l[LaForme_l.iNiveau_c - 1],
                                                      LaForme_l^.sLibelle_c, LaForme_l);
      end
      else
      begin
         aotnNoeux_l[LaForme_l.iNiveau_c - 1] := tvArbre_p.Items.AddChildObject(aotnNoeux_l[LaForme_l.iNiveau_c - 2],
                                                      LaForme_l^.sLibelle_c, LaForme_l);
      end;

      aotnNoeux_l[LaForme_l.iNiveau_c - 1].ImageIndex := TWIcone(LaForme_l^.sNonActif_c, LaForme_l^.iNiveau_c);
      aotnNoeux_l[LaForme_l.iNiveau_c - 1].SelectedIndex := TWIcone(LaForme_l^.sNonActif_c, LaForme_l^.iNiveau_c);
      iNoeud_l := LaForme_l^.iNiveau_c;
      Inc(iNbForme_l);
   end;
   tvArbre_p.Selected := tvArbre_p.Items.GetFirstNode;

   OBFormeInsee_l.Free;
   tvArbre_p.Items.EndUpdate;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 21/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TWFormateLib(sForme_p, sCodeInsee_p, sLibelle_p : string) : string;
var
   sLibelle_l : string;
begin
   if sForme_p <> '' then
   begin
      sLibelle_l := sCodeInsee_p + ' : ' + sLibelle_p + ' (' + sForme_p + ') ';
   end
   else
   begin
      sLibelle_l := sCodeInsee_p + ' : ' + sLibelle_p;
   end;
   result := sLibelle_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 22/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function  TWIcone(sNonActif_p : string; iNiveau_p : integer) : integer;
begin
   if (sNonActif_p = 'X') then
      result := 0
   else
      result := 1;//iNiveau_p;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 21/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TWVideFormes(tvArbre_p : TTreeView);
var
   tnNode_l : TTreeNode;
   LaForme_l : TPForme;
begin
   tvArbre_p.Items.BeginUpdate;
   tnNode_l := tvArbre_p.Items.GetFirstNode;

   while tnNode_l <> nil do
   begin
       LaForme_l := tnNode_l.Data;
       tnNode_l.Data := nil;
       Dispose(LaForme_l);
       tnNode_l := tnNode_l.getNext;
   end;
   tvArbre_p.Items.Clear;
   tvArbre_p.Items.EndUpdate;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 23/03/2006
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure OnExpand(tvArbre_p : TTreeView);
begin
   tvArbre_p.Items.BeginUpdate;
   tvArbre_p.FullExpand;
   tvArbre_p.Items.EndUpdate;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 23/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure OnCollapse(tvArbre_p : TTreeView);
begin
   tvArbre_p.Items.BeginUpdate;
   tvArbre_p.FullCollapse;
   tvArbre_p.Items.EndUpdate;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 15/06/2006
Modifié le ... :   /  /
Description .. : YFJ_CODEINSEE, YFJ_COOP, YFJ_FORMEGRPPRIVE, 
Suite ........ : YFJ_FORMESTE, YFJ_FORMESCI 
Mots clefs ... : 
*****************************************************************}
function GetCodeInseeFromForme(sForme_p : string;
                            var sCodeInsee_p, sFormePrive_p, sFormeSte_p,
                            sFormeSci_p, sCoop_p : string) : boolean;
var
   OBFormeInsee_l : TOB;
begin
   OBFormeInsee_l := TOB.Create('YFORMESINSEE', nil, -1);
   OBFormeInsee_l.LoadDetailDBFromSQL('YFORMESINSEE',
                                      'SELECT TOP 1 YFJ_FORME, YFJ_CODEINSEE, YFJ_COOP, YFJ_FORMEGRPPRIVE, YFJ_FORMESTE, YFJ_FORMESCI ' +
                                      'FROM YFORMESINSEE ' +
                                      'WHERE YFJ_FORME = "' + sForme_p + '" AND YFJ_NONACTIF = "-"' +
                                      'ORDER BY YFJ_FORME, YFJ_NIVEAU');

   if OBFormeInsee_l.Detail.Count > 0 then
   begin
      sCodeInsee_p  := OBFormeInsee_l.Detail[0].GetString('YFJ_CODEINSEE');
      sForme_p      := OBFormeInsee_l.Detail[0].GetString('YFJ_FORME');
      sFormePrive_p := OBFormeInsee_l.Detail[0].GetString('YFJ_FORMEGRPPRIVE');
      sFormeSte_p   := OBFormeInsee_l.Detail[0].GetString('YFJ_FORMESTE');
      sFormeSci_p   := OBFormeInsee_l.Detail[0].GetString('YFJ_FORMESCI');
      sCoop_p       := OBFormeInsee_l.Detail[0].GetString('YFJ_COOP');
   end;
   result := (OBFormeInsee_l.Detail.Count > 0);
   OBFormeInsee_l.Free;
end;

//////////////////////////////////////////////////////////////////

end.
