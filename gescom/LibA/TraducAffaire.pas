unit TraducAffaire;

                            
interface               
Uses classes,HEnt1,Paramsoc,stdctrls,ComCtrls,Controls,sysutils,HCtrls,
    grids,menus,Ent1,entgc,hmsgbox,
    HTB97,Forms,ExtCtrls ,checklst,utilgc,
{$IFDEF EAGLCLIENT}
    MenuOLX,
{$ELSE}
    MenuOLG,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB, DBCtrls,dbgrids,
{$ENDIF}
    Hout,dicoBTP,UTOB,UIUtil,licUtil,dialogs,ConfidentAffaire,UtilPgi;


// Chargement des dico ( vocabulaire , tablettes ...)
Procedure InitDicoAffaire ;
Procedure DeleteDicoAffaire;   
// Traduction des écrans
{$IFNDEF EAGLCLIENT}
Procedure TraduitAFLibGridDB(Gr : TDBGrid );
{$ENDIF}
Procedure TraduitMenu(NumMenu : integer);
procedure TraduitForm(Contain : TComponent);
Procedure TraduitEcranAffaire (F:TForm);
Procedure TraduireMetaDataAffaire; // traduction Dechamps + Detables


Procedure TraduitAFLibGridST(Gr : TStringGrid );
Function TraduitCombo (Sp : string ):string;
Function TagsInterdits :string;
// Traduction des tablettes
Function TraduitTabletteAffaire (NomTT,Lib : String): string;

Var TablettesATraduire : Tstringlist;

implementation
uses   CbpMCD,
  		CbpEnumerator;

// *****************************************************************************
// ************************* init dictionnaire  ********************************
// *****************************************************************************

{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 27/10/2000
Modifié le ... : 27/10/2000
Description .. : procédure de chargement de la traduction automatique
Mots clefs ... :
*****************************************************************}
Procedure ChargeTablettesAffaire (PrefixeTablette : string) ;
Var TobTab : TOB;
    Q : TQuery;
    i : integer;
BEGIN
TobTab := Tob.create ('liste des tablettes',Nil,-1);
Q := OpenSQL('SELECT CO_LIBELLE from COMMUN where CO_TYPE="'+PrefixeTablette+'"' ,true,-1,'',true);
if Not(Q.EOF) then TobTab.LoadDetailDB('Liste tablettes','','',Q,false);
Ferme(Q);
TablettesATraduire:=TStringList.Create;
for i := 0 to TobTab.Detail.count-1 do
    BEGIN
    TablettesATraduire.Add(TobTab.Detail[i].GetValue('CO_LIBELLE'));
    END;
TobTab.Free;
END;

Procedure InitDicoAffaire ;
Var Q : TQuery;
    PrefixeTraduc : string;
Begin
PrefixeTraduc := '';
if (Not(ctxScot in V_PGI.PGIContexte)) and (GetParamSoc('SO_AFFORMATEXER')<>'AUC') then SetParamSoc('SO_AFFORMATEXER','AUC');
if (VH_GC.AFTOBTraduction= Nil) then exit;
if ctxScot in V_PGI.PGIContexte then PrefixeTraduc := 'ATU';
if PrefixeTraduc <> '' then
    BEGIN
    Q := OpenSQL('SELECT CC_LIBELLE,CC_LIBRE from CHOIXCOD where CC_TYPE="'+PrefixeTraduc+'"' ,true,-1,'',true);
    if Not(Q.EOF) then
        VH_GC.AFTOBTraduction.LoadDetailDB('Liste traduction','','',Q,false);
    Ferme(Q);
    END;
// Gestion des tablettes à traduire
if (VH_GC.AFTOBTraduction.Detail.Count <> 0) then ChargeTablettesAffaire('ATL');
If CtxScot in V_PGi.pgicontexte then
   begin       // ilf aut traduire dans YYPARAMOBLIG le nom de la fihce affaire pour GI
   ExecuteSql ('UPDATE COMMUN SET CO_LIBRE="AFFAIRE;MISSION;1" where co_type="POB" AND CO_CODE="A08"');
   ExecuteSql ('UPDATE COMMUN SET CO_LIBRE="AFFAIRE;MISSION;1" where co_type="POB" AND CO_CODE="A10"');
   end else begin
   ExecuteSql ('UPDATE COMMUN SET CO_LIBRE="AFFAIRE;AFFAIRE;1" where co_type="POB" AND CO_CODE="A08"');
   ExecuteSql ('UPDATE COMMUN SET CO_LIBRE="AFFAIRE;AFFAIRE;1" where co_type="POB" AND CO_CODE="A10"');
   end;
// Traduction des tables (DECHAMPS) concernées
TraduireMetaDataAffaire;
End;

Procedure DeleteDicoAffaire;
BEGIN
TablettesATraduire.Free;
TablettesATraduire:=nil;
END;

// *****************************************************************************
// ********************** Traduction des écrans ********************************
// *****************************************************************************

{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 27/10/2000
Modifié le ... : 27/10/2000
Description .. : procédure de traduction automatique de tous les textes des 
Suite ........ : composants
Mots clefs ... : TRADUCTION
*****************************************************************}
// {$IFNDEF EAGLCLIENT}
procedure TraduitForm(Contain : TComponent);
var i,j,value:integer;
begin
// Traduire les conteneurs qui ont des captions à traduire avant la traduction de leur contenu
If (Contain Is TForm) Then
    BEGIN
    // Certains ecran système non traduit
    if TForm(Contain).Name = 'ParamReport' then Exit
    else if TForm(Contain).Name = 'EditEtat' then Exit
    else if TForm(Contain).Name = 'FMajStruct' then Exit
    else if TForm(Contain).Name = 'YYBANQUES' then Exit //mcd 24/06/03 avec commerciale...
    else if TForm(Contain).Name = 'FICHESIGNALE' then Exit //mcd 24/06/03 avec commerciale...
    else if TForm(Contain).Name = 'FSynthese' then Exit //mcd 24/06/03 avec commerciale...
    else if TForm(Contain).Name = 'FICHORGA' then Exit //mcd 24/06/03 avec commerciale...
    else if TForm(Contain).Name = 'ANNUAIRE' then Exit //mcd 24/06/03 avec commerciale...
    else if TForm(Contain).Name = 'YYECRVERSPIECE' then Exit //mcd 24/06/03 avec commerciale...
    else if TForm(Contain).Name = 'GCENCOURS' then Exit
    else if TForm(Contain).Name = 'GCENCOURSGC' then Exit;
    //if TForm(Contain).Name = 'FParamSoc' then Exit;
    TForm(Contain).caption :=TraduitGA(TForm(Contain).caption);
    END Else
If (Contain Is TGroupBox) Then TForm(Contain).caption :=TraduitGA(TGroupBox(Contain).caption)
Else If (Contain Is TTabSheet) Then TForm(Contain).caption :=TraduitGA(TTabSheet(Contain).caption);
//
// Boucle sur les composants contenus dans le component courant
for i:=0 to Contain.ComponentCount-1 do
begin
// Si le component courant est du type TControl
if (Contain.components[i] is TControl) then
   begin
    If (Contain.Components[i] is TButtonControl) then
        begin
        TButtonControl(Contain.Components[i]).Hint := TraduitGA(TButtonControl(Contain.Components[i]).Hint);
        if (Contain.Components[i] is TCheckBox) then
            TCheckBox(Contain.Components[i]).caption :=TraduitGA(TCheckBox(Contain.Components[i]).caption);
{$IFNDEF EAGLCLIENT}
        if (Contain.Components[i] is TDBCheckBox) then
            TCheckBox(Contain.Components[i]).caption :=TraduitGA(TCheckBox(Contain.Components[i]).caption);
{$ENDIF}
        if (Contain.Components[i] is TRadioButton) then
            TRadioButton(Contain.Components[i]).caption :=TraduitGA(TRadioButton(Contain.Components[i]).caption);

        end Else
    If (Contain.Components[i] is TCustomEdit) then
        begin
        TCustomEdit(Contain.Components[i]).Hint := TraduitGA(TCustomEdit(Contain.Components[i]).Hint);
        end Else
    If (Contain.Components[i] is TLabel) then
        begin
            TLabel(Contain.Components[i]).caption :=TraduitGA(TLabel(Contain.Components[i]).caption);
        TLabel(Contain.Components[i]).Hint := TraduitGA(TLabel(Contain.Components[i]).Hint);
        end Else
    If (Contain.Components[i] is TStaticText) then
        begin
        TStaticText(Contain.Components[i]).caption :=TraduitGA(TStaticText(Contain.Components[i]).caption);
        TStaticText(Contain.Components[i]).Hint :=TraduitGA(TStaticText(Contain.Components[i]).Hint);
        end Else
    If (Contain.Components[i] is TCheckListBox) then
        begin
        TChecklistBox(Contain.Components[i]).items.Commatext :=TraduitGA(TChecklistBox(Contain.Components[i]).items.Commatext);
        TChecklistBox(Contain.Components[i]).Hint :=TraduitGA(TChecklistBox(Contain.Components[i]).Hint);
        end Else
    if (Contain.Components[i] is TListBox) then
        begin
        TlistBox(Contain.Components[i]).items.Commatext :=TraduitGA(TlistBox(Contain.Components[i]).items.Commatext);
        TlistBox(Contain.Components[i]).Hint :=TraduitGA(TlistBox(Contain.Components[i]).Hint);
        end Else
    if (Contain.Components[i] is TComboBox) then
        Begin
        TComboBox(Contain.Components[i]).Hint := TraduitGA(TComboBox(Contain.Components[i]).Hint);
        Value := TComboBox(Contain.Components[i]).itemindex;
        TComboBox(Contain.Components[i]).items.Commatext :=TraduitGA(TComboBox(Contain.Components[i]).items.Commatext);
        TComboBox(Contain.Components[i]).itemindex := Value;
        End Else
    If (Contain.Components[i] is TGraphicControl) then
        begin
        TGraphicControl(Contain.Components[i]).Hint := TraduitGA(TGraphicControl(Contain.Components[i]).Hint);
        If (Contain.Components[i] is TToolBarButton97) then
            begin
            TToolBarButton97(Contain.Components[i]).Caption :=TraduitGA(TToolBarButton97(Contain.Components[i]).Caption);
            end;
        end Else
    if (Contain.Components[i] is TStringGrid) then
        TraduitAFLibGridST(TStringGrid(Contain.Components[i])) Else
{$IFNDEF EAGLCLIENT}
    if (Contain.Components[i] is TDBGrid) then
        TraduitAFLibGridDB(TDBGrid(Contain.Components[i])) Else
{$ENDIF}
   // Appel récursif de la fonction de traduction pour les containers
   If (Contain.Components[i] Is TForm) or (Contain.Components[i] Is TTabSheet) or
      (Contain.Components[i] Is TPageControl) or (Contain.Components[i] Is TgroupBox) or
      (Contain.Components[i] Is Tpanel) Then TraduitForm(Contain.Components[i]);
   end
else
// Si le component courant est du type TMenu
if (Contain.components[i] is TMenu) then
   begin
   if (Contain.components[i] is Tpopupmenu) then
      for j:=0 to Tpopupmenu(Contain.components[i]).Items.Count-1 do
         TMenuItem(Tpopupmenu(Contain.components[i]).items[j]).Caption :=TraduitGA(TMenuItem(Tpopupmenu(Contain.components[i]).items[j]).Caption);
   end else
// Si le component courant est du type THMsgBox
if (Contain.components[i] is THMsgBox) then
   begin
   if (Contain.components[i] is THMsgBox) then
      THMsgBox(Contain.components[i]).Mess.Commatext :=TraduitGA(THMsgBox(Contain.components[i]).Mess.Commatext)
   end;
end;
end;

Procedure TraduitEcranAffaire (F:TForm);
BEGIN
TraduitForm(F);
END;

// {$ENDIF}

{$IFNDEF EAGLCLIENT}
Procedure TraduitAFLibGridDB(Gr : TDBGrid );
Var i : integer;
Begin
If Not(Gr is TDBGrid) Then Exit;
If Not(ctxScot in V_PGI.PGIContexte) Then  Exit;
if (Gr.Columns.Count = 1) then Exit ;
For i:=0 to Gr.Columns.Count-1 do
    Begin
    if Gr.columns[i].Field <> Nil then
        Gr.columns[i].Field.DisplayLabel:=TraduitGA(Gr.columns[i].Field.DisplayLabel);
    End;
End;
{$ENDIF}

Procedure TraduitAFLibGridST(Gr : TStringGrid );
Var i : integer;
Begin
If Not(Gr is TStringGrid) Then Exit;
If Not(ctxScot in V_PGI.PGIContexte) Then  Exit;
if (Gr.ColCount = 1) then Exit ;
if (Gr.FixedRows = 0) then Exit ;
For i:=0 to Gr.ColCount-1 do
    Begin
    Gr.Cells[i, 0]:=TraduitGA(Gr.Cells[i, 0]);
    End;
End;


// *****************************************************************************
// ********************** Traduction des tablettes *****************************
// *****************************************************************************

{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 29/08/2000
Modifié le ... : 29/08/2000
Description .. : Traduction des contenus des tablettes / dico affaire.
Suite ........ : Seules certaines tablettes sont traduites
Mots clefs ... : TRADUCTION;TABLETTE
*****************************************************************}
Function TraduitTabletteAffaire (NomTT,Lib : String): string;
Var NumTab,ii : integer ;
BEGIN
result:=Lib;
if (TablettesATraduire=nil) then exit;
NumTab := TablettesATraduire.indexof(NomTT);
if NumTab < 0 then result := Lib
              else begin
                        // mcd 24/10/02 traduit zone partie X code affaire
                   ii:= pos('PARTIE',AnsiUppercase(lib));
                   if ( ii >0 )then
                      begin
                      if  copy (lib,ii+7,1) ='1' then Result:=VH_GC.CleAffaire.Co1Lib
                      else  if  copy (lib,ii+7,1) ='2' then if VH_GC.CleAffaire.Co2Visible then Result:=VH_GC.CleAffaire.Co2Lib
                      else  if  copy (lib,ii+7,1) ='3' then if VH_GC.CleAffaire.Co3Visible then Result:=VH_GC.CleAffaire.Co3Lib;
                      end
                   else  result := traduitGA (Lib);
                   end;
END;


{**********************************************************
Auteur  ...... : MC DESSEIGNET
Créé le ...... : 29/05/2001
Modifié le ... :   /  /    
Description .. :traduit pour les états les combos et champ libre
reçoit le nom de la zone
Mots clefs ... :
*****************************************************************}
Function TraduitCombo (sp : String): string;
Var     titre:string;
 //   Typetab, typechamp,Num,Name,Titre:string;

BEGIN
     // permet pour les états, la traduction des champs libres.
     // fait à partir de traducchamplibre  onargument
result:='';
GCTitreZoneLibre (Sp, Titre);
if (titre='') then begin
              // on regarde si famille article
   if sp='GA_FAMILLENIV1' then Titre := rechdom ('GCLIBFAMILLE','LF1',False)
   else if sp='GA_FAMILLENIV2' then Titre := rechdom ('GCLIBFAMILLE','LF2',False)
   else if sp='GA_FAMILLENIV3' then Titre := rechdom ('GCLIBFAMILLE','LF3',False)
   else if sp='ATBXFAMILLENIV1' then Titre := rechdom ('GCLIBFAMILLE','LF1',False)
   else if sp='ATBXFAMILLENIV2' then Titre := rechdom ('GCLIBFAMILLE','LF2',False)
   else if sp='ATBXFAMILLENIV3' then Titre := rechdom ('GCLIBFAMILLE','LF3',False);
   end;
result:=titre;
END;


// *****************************************************************************
// ********************** Traduction des menus *********************************
// *****************************************************************************
{$IFNDEF EAGLCLIENT}
Function GetLibItem(nTag : integer) : String ;
var TreeN : TTreeNode ;
BEGIN
result := '';
TreeN:=GetTreeItem(FMenuG.Treeview,nTag) ;
if TreeN<>Nil then result := TreeN.Text ;
END;
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :
Modifié le ... :
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TraduitMenu(NumMenu : integer);
Var ListeItems,  tmp: String;
    ii: integer;
    TobPlanning :TOB;
    vBoPlanDeCharge : Boolean;
    Q : Tquery;
    St : STring;

BEGIN                 

  Listeitems := '';
  // ************** Traduction du menu *******************
  Case NumMenu of

    144,74: begin

              // menu supprimer en cwas
              {$IFDEF EAGLCLIENT}
              Fmenug.removeGroup(74100,true); //menu socité (paramètre, parpaiece,ETB,Cpteur
              Fmenug.removeItem(74151); //Regime fiscaux cacher dans un 1er temps car pas en GC
              Fmenug.removeItem(74152); // TVA cacher dans un 1er temps car pas en GC
              Fmenug.removeItem(74155); // Devise cacher dans un 1er temps car pas en GC
              Fmenug.removeItem(74159);  // Region cacher dans un 1er temps car pas en GC
              Fmenug.removeItem(74160);  //code postaux
              Fmenug.removeItem(74205);  //méthode arrondi
              Fmenug.removeItem(74206);  //venitl anamlytique
                              
              Fmenug.removeGroup(74250,true); //menu paramètre etat
              Fmenug.removeItem(74701);  //utilisatuer
              Fmenug.removeItem(74702);  //groupe util
              Fmenug.removeItem(74703);  //utilisatuer connecté
              Fmenug.removeItem(74704);  //suivi activite
              Fmenug.removeItem(74705);  //raz connection

              Fmenug.removeItem(74751);  //saisie blocage mission
              Fmenug.removeItem(74802);  //recopie de société
 
              Fmenug.removeGroup(74650,true); //import/export
// quand eactivite sera eagl
{              Fmenug.removeItem(74656);  //groupe util
              Fmenug.removeItem(74651);  //utilisatuer connecté
              Fmenug.removeItem(74652);  //suivi activite
              Fmenug.removeItem(74654);  //raz connection
}
              Fmenug.removeGroup(74600,true); //menu utilitaire (sat ...)
              Fmenug.removeGroup(74800,true); //menu recopie sté
             {$ENDIF}
               
              // on met les libellés paramétrables dans le menu
              For ii:=1 to 9 do
              begin
								tmp := RechDomZoneLibre ('CT'+IntToStr(ii), False);
//                tmp:= RechDom('GCZONELIBRE','CT'+IntToStr(ii),FALSE);
                if tmp ='.-' then
                  FMenug.removeItem(74410+ii)
                else
                  FMEnuG.RenameItem (74410+ii,tmp);
								tmp := RechDomZoneLibre ('ET'+IntToStr(ii), False);
//                tmp:=RechDom('GCZONELIBRE','ET'+IntToStr(ii),FALSE);
                if tmp ='.-' then
                  FMenug.removeItem(74469+ii)
                else
                  FMEnuG.RenameItem (74469+ii,tmp);
								tmp := RechDomZoneLibre ('AT'+IntToStr(ii), False);
//                tmp:=RechDom('GCZONELIBRE','AT'+IntToStr(ii),FALSE);
                if tmp ='.-' then
                  FMenug.removeItem(74440+ii)
                else
                  FMEnuG.RenameItem (74440+ii,tmp);
								tmp := RechDomZoneLibre ('MT'+IntToStr(ii), False);
//                tmp:=RechDom('GCZONELIBRE','MT'+IntToStr(ii),FALSE);
                if tmp ='.-' then
                  FMenug.removeItem(74460+ii)
                else
                  FMEnuG.RenameItem (74460+ii,tmp);
								tmp := RechDomZoneLibre ('RT'+IntToStr(ii), False);
//                tmp:=RechDom('GCZONELIBRE','RT'+IntToStr(ii),FALSE);
                if tmp ='.-' then
                  FMenug.removeItem(74430+ii)
                else
                  FMEnuG.RenameItem (74430+ii,tmp);
								tmp := RechDomZoneLibre ('TT'+IntToStr(ii), False);
//                tmp:=RechDom('GCZONELIBRE','TT'+IntToStr(ii),FALSE);
                if tmp ='.-' then
                  FMenug.removeItem(74580+ii)
                else
                  FMEnuG.RenameItem (74580+ii,tmp);
								tmp := RechDomZoneLibre ('BT'+IntToStr(ii), False);
//                tmp:=RechDom('GCZONELIBRE','BT'+IntToStr(ii),FALSE);
                if tmp ='.-' then FMenug.removeItem(74540+ii)
                else FMEnuG.RenameItem (74540+ii,tmp);
              end;

              For ii:=1 to 3 do
              begin
								tmp := RechDomZoneLibre ('FT'+IntToStr(ii), False);
//                tmp:=RechDom('GCZONELIBRE','FT'+IntToStr(ii),FALSE);
                if tmp ='.-' then FMenug.removeItem(74560+ii)
                else FMEnuG.RenameItem (74560+ii,tmp);
  								tmp := RechDomZoneLibre ('PT'+IntToStr(ii), False);
//              tmp:=RechDom('GCZONELIBRE','PT'+IntToStr(ii),FALSE);
                if tmp ='.-' then FMenug.removeItem(74450+ii)
                else FMEnuG.RenameItem (74450+ii,tmp);
              end;

								tmp := RechDomZoneLibre ('CTA', False);
//              tmp:=RechDom('GCZONELIBRE','CTA',FALSE);
              if tmp ='.-' then FMenug.removeItem(74420)
              else FMEnuG.RenameItem (74420,tmp);

//              tmp:=RechDom('GCZONELIBRE','ETA',FALSE);
								tmp := RechDomZoneLibre ('ETA', False);
              if tmp ='.-' then FMenug.removeItem(74479)
              else FMEnuG.RenameItem (74479,tmp);

//              tmp:=RechDom('GCZONELIBRE','ATA',FALSE);
              tmp := RechDomZoneLibre ('ATA', False);
              if tmp ='.-' then FMenug.removeItem(74450)
              else FMEnuG.RenameItem (74450,tmp);

//              tmp:=RechDom('GCZONELIBRE','BTA',FALSE);
              tmp := RechDomZoneLibre ('BTA', False);
            if tmp ='.-' then FMenug.removeItem(74560)
              else FMEnuG.RenameItem (74560,tmp);

//              tmp:=RechDom('GCZONELIBRE','MTA',FALSE);
              tmp := RechDomZoneLibre ('MTA', False);
              if tmp ='.-' then FMenug.removeItem(74489)
              else FMEnuG.RenameItem (74489,tmp);

//              tmp:=RechDom('GCZONELIBRE','RTA',FALSE);
              tmp := RechDomZoneLibre ('RTA', False);
              if tmp ='.-' then FMenug.removeItem(74499)
              else FMEnuG.RenameItem (74499,tmp);

//              tmp:=RechDom('GCZONELIBRE','TTA',FALSE);
              tmp := RechDomZoneLibre ('TTA', False);
              if tmp ='.-' then FMenug.removeItem(74590)
              else FMEnuG.RenameItem (74590,tmp);
              FMenug.removeItem(74457); // C.B 17/01/03 libellés libres pièces / dates pas gérées
             {  remis le 31/10/2003 car utilisé en clientèle par Franck Wantiez
             if ctxtempo in V_PGI.PGIContexte then
              begin
                FMEnuG.removeItem (74253);  // edition word Proposition
                FMEnuG.removeItem (74263);  // edition word Affaire
              end;
              }
            end;
      92: begin
            if ctxscot in V_PGI.PGIContexte then FMEnuG.RenameItem (92102,'Contacts');
          end;
  End;

  // La liste des items contenant des champs (affaires, ressources ...) est conservée
  // pour la traduction auto si les libellés sont paramétrables non géré actuellement
  {** Item:=(Trim(ReadTokenSt(ListeItems)));
  While (Item <>'') do
      BEGIN
      NumItem := strToInt(Item);
      Lib := TraduitGA(GetLibItem(NumItem));
      FMenuG.RenameItem(NumItem,Lib) ;
      Item:=(Trim(ReadTokenSt(ListeItems)));
      END;  **}

  //////// *********** Gestion des différences entre les menus *******************
  //******* Menu affaire de base *************
  if (numMenu = 71) or (numMenu = 141) then
  Begin
    if not(GetParamSoc('SO_PgLienRessource')) or (GetParamSoc('SO_AFLIENPAIEDEC')) then
      FMenug.removeItem(71445);  //pas le param lien ressource salairé ou paie décentralisée
    if (ctxScot In V_PGI.PGIContexte) then
    Begin
       FMenug.RemoveItem (71205);  //suivi prop à suurpierm
      if Not(VH_GC.AFGestionCom) then FMenug.removeItem(71501); // pas gestion commercial/apporteur
      if (GetParamSoc('SO_AfFormatExer') = 'AUC') then FMenug.removeItem(71408)  //duplication N en N+1
         else if (V_PGI.PassWord <> CryptageSt(DayPass(Date)))then  FMenug.removeItem(71412);    //epuration ech facturation
      If not(GetParamSOc('SO_AFGESTIONTARIF'))then
        begin
        Fmenug.removeGroup(-71600,true);  //pas de menu tarif en GI
(*          begin
          FMenug.removeItem(71601); // pas tarif
          FMenug.removeItem(71602); // pas tarif
          FMenug.removeItem(71603); // pas tarif
          FMenug.removeItem(71604); // pas tarif
          FMenug.removeItem(71605); // pas tarif
          FMenug.removeItem(71606); // pas tarif
          FMenug.removeItem(71615); // pas tarif
          FMenug.removeItem(71616); // pas tarif
          FMenug.removeItem(71611); // pas tarif TTC
          FMenug.removeItem(71612); // pas tarif TTC
          FMenug.removeItem(71613); // pas tarif TTC
          FMenug.removeItem(-71610); // pas tarif TTC
          end; *)
        end;
     if VH_GC.GAPlanningSeria = False then
        begin
         FMenuG.RemoveItem(-71260);  //impression tache
        end;
    end;

    {$IFNDEF DEBUG}
    FMenug.RemoveItem (71105);  // Mul Tiers avec critères Annuaire/DPs $$$JP
    {$ENDIF}

    if Not(VH_GC.AFGestionCom) then FMenuG.removeItem (71501); // sup apporteurs
    if Not(VH_GC.CleAffaire.GestionAvenant) then Fmenug.removeItem(-71110);
    if (ctxScot In V_PGI.PGIContexte) and ((VH_GC.AFFormatExer='AA') or  (VH_GC.AFFormatExer='AM')or (VH_GC.AFFormatExer='A')) then
       Fmenug.removeItem(71413);

    {$IFDEF SANSCOMPTA} // fonction ouverture fermeture de compte tiers
       Fmenug.removeItem(71404); Fmenug.removeItem(71405); Fmenug.removeItem(71406);
    {$ENDIF}

    {$IFDEF EAGLCLIENT}  // non dispo EAGL
    // modif en série client,                article                     affaire
    Fmenug.removeItem(71403); Fmenug.removeItem(71410);
    //Fmenug.removeItem(71415);  mcd 22/10/2002
    Fmenug.removeGroup(-71415,true);  //pas de mofi par lot mission
    //Fmenug.removeGroup(-71441,true);  //pas de mofi par lot ressource
    FMenuG.RemoveGroup(71430,True); // plus d'option dans le groupe traitement tiers !!!
    FmenuG.RemoveItem(71607); // mise à jour tarifaire
    FmenuG.RemoveItem(71407); // Suppr arti
    {$ENDIF}
  End

  else

  //******* Menu paramétrage *************
  if (numMenu = 74) or (numMenu = 144) then
      BEGIN
      if GetParamSoc('SO_GCPIECEADRESSE') then FMEnuG.removeItem (74614);//bascule adresse
      FMenuG.RemoveItem(-74712); // Tablette hirarchqiue pas OK V5 06/2003
      if (ctxScot In V_PGI.PGIContexte) then
          BEGIN
          if (GetParamSoc('SO_AfFormatExer') <> 'AUC') then FMenuG.RemoveItem(74109) // pas assistant codif affaire
               else FMenuG.RemoveItem(74602); // pas reprise SCOT pour l'instant
          if Not(V_PGI.SAV) then
              BEGIN
              FMenuG.RemoveItem(74111); // param listes de saisies
              FMenuG.RemoveItem(74651); // etiers
              FMenuG.RemoveItem(74656); // eaffaire
              END;
          if Not (V_PGI.PassWord = CryptageSt(DayPass(Date))) then  // Mot de passe du jour
              begin
              FMenuG.RemoveItem(74108); // assistant de création PCL
              FMenuG.RemoveItem(74613); // Modif Analytique
              FMenuG.RemoveItem(74618); // impression mcd
              FMenuG.RemoveItem(-74610); // outil assistance               
              end;
                  //cas dossier GI : toujours avec lanceur :  on n'a pas non plus les utilisateurs et les fct
                    // sauf en mode DEBUG pour test interne sans lanceur
          if (GetParamsoc('SO_AFLIENDP') =true) then
           begin
               // si pas lien DP, il faut quand même saisir la confidentialité
             FMenuG.RemoveItem(74702); // groupe utilisateur
           end
          else
            begin
             FMenuG.RemoveItem(-74616); // Menu global des modifs sur annuaire PL le 18/11/02
             FMenuG.RemoveItem(74604); // ctrl annuaire/tiers
             FMenuG.RemoveItem(74605); // lien annuaire/tiers rapide
             FMenuG.RemoveItem(74607); // Doublons annauire
             FMenuG.RemoveItem(74606); // doublons annuaire/tiers rapide
            end;
  {$ifndef DEBUG}
          FMenuG.RemoveItem(74701); //  utilisateur
          FMenuG.RemoveItem(74703); //  utilisateur  connectés
          FMenuG.RemoveItem(74704); //  suivi d'activité
          FMenuG.RemoveItem(74705); //  raz connection
  {$endif}
          FMenuG.RemoveItem(74457); //  Dates libre spièces non géré pour l'instant
          FMenuG.RemoveItem(-74520); //  approteur table libre
          FMenuG.RemoveItem(74522); //  approteur table libre
          FMenuG.RemoveItem(74523); //  apporteur date libre
          if (GetParamSoc('SO_AFAPPAVECBM')<>True) then
             begin
            FMenuG.RemoveItem(73605); //  paramétrage boni/mali
            FMenuG.RemoveItem(74207); //  justificatif boni/mali
            end;

          END; // fin uniquement Gestion interne
     If (( not V_PGI.Superviseur) and ( not V_PGI.Controleur)) or (GetParamsoc('SO_AFTYPECONF')<>'AGR') then
         begin
          FMenuG.RemoveItem(-74706); //  Menu Confidentailité
          FMenuG.RemoveItem(74706); //  Confidentailité
          FMenuG.RemoveItem(74707); //  Groupe de confidentialité
         end;
     if VH_GC.GAPlanningSeria = False then
        begin
         FMenuG.RemoveItem(-74570);  //table libre tache
         FMenuG.RemoveItem(-74580);
        end;
     if VH_GC.GAAchatSeria = False then
        begin
        FMenug.removeItem(-74560);   // table libre fournissuer
        FMenug.removeitem(-74550);   // table libre fournissuer
        end;
     if VH_GC.GASaisieDecSeria = False then
        begin
        FMenug.removeItem(74655);   // export fichier
        FMenug.removeItem(74654);   // Saisie decentralise
        end;
     if Not(VH_GC.GAPlanningSeria) then
        begin
        FMenug.removeitem(-74380);   // Famille des tâches
        FMenug.removeitem(-74570);   // libellés des libres tâches
        FMenug.removeitem(-74580);   // Libres tâches
        end;
     if (GetParamSoc('SO_AFAPPCUTOFF')<>True) then
     // PL le 05/11/02 : ouverture à Scot
     //or ((ctxScot In V_PGI.PGIContexte) and Not (V_PGI.PassWord = CryptageSt(DayPass(Date)))) then
          // Si le paramsoc des gestion des cutoff n'est pas selectionne ou
          // si on est en Scot et pas entre avec le mot de passe du jour,
          // on ne peut pas modifier la formule des cutoff
          FMenuG.RemoveItem(74752);   // pas de  formule cutoff
     if CtxScot in V_PGI.PgiCOntexte then FMenuG.RemoveItem(74753);   // lien paie
     if not GetParamSoc('SO_AFLIENPAIEANA') and not GetPAramSoc('SO_AFLIENPAIEVAR') then
       FMenuG.RemoveItem(74753);
     if  not GetParamSoc('SO_AFVARIABLES') then FMenuG.RemoveItem(74754);
     END else
  //******* Menu Activité  *************
  if (numMenu = 72) or (numMenu = 142) then
      BEGIN
        // quand remis, les cacher si so_aflinepaievar est faux !!!
      if CtxScot in V_PGI.PgiCOntexte then FMenuG.RemoveItem(72510);   // pas de lien paie
      if CtxScot in V_PGI.PgiCOntexte then FMenuG.Removeitem(72511);   // pas de lien paie
      if (GetParamSoc('SO_AFVISAACTIVITE')<>True) then FMenuG.RemoveItem(72501); //pas de visa
      if (GetParamSoc('SO_AFAPPCUTOFF')<>True) then FMenuG.Removegroup(-72700,True);   // pas de  cutoff
      If (GetParamSoc('SO_AFCUTMODEECLAT')='GLO') then FmenuG.RemoveItem(72703); // modification détaillée
      FMenuG.RemoveGroup(72300,True);   // pas de gestion histo dans un 1er temps
      if CtxScot in V_PGI.PgiCOntexte then  FMenuG.RemoveItem(72706);  // pas d'etat cut off sur 12 mos
      If Not(SaisieActiviteManager) then
         begin // mcd 14/10/2002 cache cube si pas saisie manager
         FMenuG.RemoveItem(72620);
         FMenuG.RemoveItem(-72110);
         end;
      FMenuG.RemoveItem(72707); //Génération en comptabilité pas diffusé en 5.0
      if not GetParamSoc('SO_AFLIENPAIEANA') and not GetPAramSoc('SO_AFLIENPAIEVAR') then
      begin
         FMenuG.RemoveItem(72510);
         FMenuG.RemoveItem(72511);
      end;
    END else
   //******* Menu Planning*************
    // C.B 24/05/02
    if (numMenu = 153) or  (numMenu = 154) then
      Begin

        // suppression du menu editions
        // pas encore d'éditions
        FMenuG.RemoveGroup(153600, True);
                                          
        vBoPlanDeCharge := GetParamSoc('SO_AFPLANDECHARGE',true);
        St := 'SELECT APP_CODEPARAM, APP_LIBELLEPARAM FROM AFPLANNINGPARAM ';
        St := St + 'WHERE RIGHT(APP_CODEPARAM, 1) = "J"';

        Q := OpenSQL(St, True,-1,'',true);
        TobPlanning := TOB.Create ('AFPLANNINGPARAM', Nil, -1);

        Try
          TobPlanning.LoadDetailDB ('AFPLANNINGPARAM', '', '',Q, True);

          FMenuG.RemoveItem(153304); // revalorisation pas encore gérée
          if vBoPlanDeCharge then
            begin
              FMenuG.RemoveItem(153103); // raf en quantité
              FMenuG.RemoveItem(153104); // raf en montant
              FMenuG.RemoveItem(153204);
              FMenuG.RemoveItem(153203);
              FMenuG.RemoveItem(153202);
              FMenuG.RemoveItem(153201);
              FMenuG.RemoveItem(153301); // génération du planning
              FMenuG.RemoveItem(153302); // génération de l'activité
              FMenuG.RemoveGroup(153400, True); // Paramètres
              //analyse planning jour
              FMenuG.RemoveItem(153503);
              FMenuG.RemoveItem(153504);
              FMenuG.RemoveItem(-153505);
              FMenuG.RemoveItem(153508);
              //FMenuG.RemoveGroup(153500, True); // Analyse
            end
          else
            begin
              FMenuG.RemoveItem(153205); // plan de charge
              FMenuG.RemoveItem(153206); // plan de charge en montant
              // analyse plan de charge
              FMenuG.RemoveItem(153501); // tobviewer
              FMenuG.RemoveItem(153502); // tobviewer ressource

              if not getParamSoc('SO_AFGESTIONRAF') then
                begin
                  FMenuG.RemoveItem(153103); // raf en quantité
                  FMenuG.RemoveItem(153104); // raf en montant
                end;

              // planning 4
              if (TobPlanning.Detail.count < 4) then
                FMenuG.RemoveItem(153204)
              else
                FMenuG.RenameItem(153204, TobPlanning.Detail[3].GetValue('APP_LIBELLEPARAM'));

              // planning 3
              if (TobPlanning.Detail.count < 3) then
                FMenuG.RemoveItem(153203)
              else
                FMenuG.RenameItem(153203, TobPlanning.Detail[2].GetValue('APP_LIBELLEPARAM'));

              // planning 2
              if (TobPlanning.Detail.count < 2) then
                FMenuG.RemoveItem(153202)
              else
                FMenuG.RenameItem(153202, TobPlanning.Detail[1].GetValue('APP_LIBELLEPARAM'));

              // planning 1
              if (TobPlanning.Detail.count < 1) then
                FMenuG.RemoveItem(153201)
              else
                FMenuG.RenameItem(153201, TobPlanning.Detail[0].GetValue('APP_LIBELLEPARAM'));
            end;

        finally
          TobPlanning.Free;
          Ferme(Q);
        end;

      end  else

  //******* Menu Achats*************
  if (numMenu = 138) or (numMenu = 152) then
  Begin
    //    if (GetParamSoc('SO_AFAPPCUTOFF')<>True)  then
    FMenuG.RemoveGroup(138700,True);   // Cutoff fournisseurs
    if (GetParamSoc('SO_AFCUTMODEECLAT')='GLO')  then FMenuG.RemoveItem(138704); // modif détaillé

    {$ifndef DEBUG}
    FMenuG.RemoveItem(138130);   // C.B historisation pas livrée pour l'instant
    FMenuG.RemoveItem(138140);   // C.B épuration pas livrée pour l'instant
    {$ENDIF}
    {$IFDEF CCS3}
    FMenuG.RemoveItem(138301);   // visa des piece
    FMenuG.RemoveGroup(138400,true);   // visatarif fournisseur
    {$ENDIF}

  End

  else

    //******* Menu Planning*************
  {if (numMenu = 153) or (numMenu = 154) then
      BEGIN
      FMenuG.RemoveItem(153202);   // tache par ress/affaire
      FMenuG.RemoveItem(153203);   // ressource par aff/tache
      FMenuG.RemoveItem(153204);   // taux occupation
      FMenuG.RemoveItem(153206);   //frais
      FMenuG.RemoveGroup(153300,True);   // traitement
      end else
  }  //******* Menu Analyse*************
  if (numMenu = 139) or (numMenu = 151) then
      BEGIN
      if (VH_GC.GAAchatSeria = False) And (ctxScot in V_PGI.PGIContexte) then
        begin
        FMenuG.RemoveGroup(139200,True);   // analyse achats
        end;
      If Not(SaisieActiviteManager) then
         begin // mcd 14/10/2002 cache accès consultation si pas saisie manager
         FMenuG.RemoveGroup(139300,True);   // analyse
         end;
//      {$ifdef EAGLCLIENT}   // remis car le 5eme argument ne passe pas en e-agl
//                              a priori ok le 21/10/03
//       FMenuG.RemoveItem(139211);  // graph
//       FMenuG.RemoveItem(139221);  // graph
//       FMenuG.RemoveItem(73531);   // Graph
//       FMenuG.RemoveItem(73541);   // Graph
//     {$endif}

      end   else

    //******* Menu Facturation *************
  if (numMenu = 73) or (numMenu = 143) then
  BEGIN
// cInClientOnyx then  suppression des tests clt onyx, conditionné avec revision en attendant
// que ce soit généralisé à tout le monde
   if not GetParamSoc('SO_AFREVISIONPRIX') then
    FMenuG.RemoveItem(73408);   // Gen avoir globaux pas ok au 01/10/2003.. enlever

    // révision de prix
    if not GetParamSoc('SO_AFREVISIONPRIX') then
    begin
      FMenuG.RemoveItem(73409);
      FMenuG.RemoveItem(73203);
    end;
 
  If (GetParamSoc('SO_AFFACTPARRES') = 'SAN') then FMenuG.RemoveGroup(73500,True);   // fcature éclatée
      //C.B 14/11/02 vieux plus utilisé
      //if Not(V_PGI.SAV) and not(ctxScot In V_PGI.PGIContexte) then
      FMenuG.RemoveItem(73405); // cde => Fac

  {$ifndef DEBUG}
      FMenuG.RemoveItem(73120);   // C.B historisation pas livré pour l'instant
      FMenuG.RemoveItem(73121);   // C.B historisation pas livré pour l'instant
      FMenuG.RemoveItem(73402);   // Test en debug
      FMenuG.RemoveItem(73618);   // validation def  Redondant avec celui autre menu
  {$ENDIF}

      // C.B 04/02/03
       // mcd 17/04/03 if GetInfoParPiece('FPR','GPP_VISA')='-' then
      if (GetInfoParPiece('FPR','GPP_VISA')='-') and (GetInfoParPiece('APR','GPP_VISA')='-')  then
        FMenuG.RemoveItem(73406);

      // on l'ote définitivement
      if Not(VH_GC.AFGestionCom) then FMenuG.removeItem (73306); // sup commissionnement

      if (GetParamSoc('SO_AFGESTIONAPPRECIATION')<>True) then
          FMenuG.RemoveGroup(73600,True);   // apprec
      if CtxScot in V_PGI.PgiCOntexte  then FMenuG.RemoveItem(73601); // Modif Boni/mali pas OK V5.0
      if CtxTempo in V_PGI.PgiCOntexte  then
        begin
        FMenuG.RemoveItem(73308);   // Etat prepa factuartion GI
        FMenuG.RemoveGroup(73500,True);   // fcature éclatée
        end;
  END else
  if (numMenu= 140) then
     //**** Administration ****
     BEGIN
       if GetParamSoc('So_GCPIECEADRESSE') then FMEnuG.removeItem (74614);//bascule adresse

       // menu supprimer en cwas
       {$IFDEF EAGLCLIENT}
       Fmenug.removeGroup(74650,true); //import/export
// quand eactivite sera eagl
{       Fmenug.removeItem(74656);  //groupe util
       Fmenug.removeItem(74651);  //utilisatuer connecté
       Fmenug.removeItem(74652);  //suivi activite
       Fmenug.removeItem(74654);  //raz connection
}
       Fmenug.removeGroup(74600,true); //menu utilitaire (sat ...)
       Fmenug.removeGroup(74800,true); //menu recopie sté
       Fmenug.removeItem(74701);  //utilisatuer
       Fmenug.removeItem(74702);  //groupe util
       Fmenug.removeItem(74703);  //utilisatuer connecté
       Fmenug.removeItem(74704);  //suivi activite
       Fmenug.removeItem(74705);  //raz connection
       Fmenug.removeItem(74620);  //Restriction utilisateurs
       {$ENDIF}

       FMenuG.RemoveItem(-74712); // Tablette hirarchqiue pas OK V5 06/2003
       If (( not V_PGI.Superviseur) and ( not V_PGI.Controleur)) or (GetParamsoc('SO_AFTYPECONF')<>'AGR') then
         begin
          FMenuG.RemoveItem(-74706); //  Menu Confidentailité
          FMenuG.RemoveItem(74706); //  Confidentailité
          FMenuG.RemoveItem(74707); //  Groupe de confidentialité
         end;
   		if VH_GC.GASaisieDecSeria = False then
        begin
        	FMenug.removeItem(74655);   // export fichier
        	FMenug.removeItem(74654);   // Saisie decentralise
        end;
     FMenuG.RemoveItem(74652);   // e-commande (jamais utilisé...)
     if Not(V_PGI.SAV) then
    // mcd 25/06/02 if Not (V_PGI.PassWord = CryptageSt(DayPass(Date))) then  // Mot de passe du jour
        begin
          FMenuG.RemoveItem(74656); FMenuG.RemoveItem(74651);  //import e-Tiers +  Affaires
        end;
     if Not (V_PGI.PassWord = CryptageSt(DayPass(Date))) then
        begin
         FMenuG.RemoveItem(74613); // modif analytique
         FMenuG.RemoveItem(74699); // util de dbeug ???
         FMenuG.RemoveItem(74618); // impression mcd
         FMenuG.RemoveItem(-74610); // outil assistance
         end;
     {$IFDEF CCS3}
      FMenuG.RemoveItem(74708); // zones obligatoires
      FMenuG.RemoveItem(74709); // restrictions fiches
      FMenuG.RemoveGroup(-74712,true); // tablettes hierarchiques
      FMenuG.RemoveGroup(74650,true); // import/export
     {$ENDIF}
     END;

  // pas encore developpé GA : revision de prix
  if (numMenu = 262) then
  begin
//   FMenuG.RemoveItem(262401);
   FMenuG.RemoveItem(262504);
  end;

  // traitement seul des factures de regul supprimé pour l'instant
  FMenuG.RemoveItem(262404);

END;

Function TagsInterdits :string;
//Var    TobModelePlanning:TOB;
begin
// Favoris
// fct qui met tous les tags interdit en GI et/ou GA

  If Not(SaisieActiviteManager) then result:=Result+'72620;139300;-72110;';             // mcd 14/10/2002
  if GetParamSoc('So_GCPIECEADRESSE') then Result :=Result+'74614;';//bascule adresse
  if not GetParamSoc('SO_AFREVISIONPRIX') then Result :=Result + '73409;73203;262504;';
  if (ctxtempo In V_PGI.PGIContexte) then  REsult := result+'73308;73500;';
  if (ctxScot In V_PGI.PGIContexte) then
    BEGIN
    Result := Result + '71205;';
    Result := Result + '92102;';
    if Not(VH_GC.AFGestionCom) then result:=result+'71501;'; // pas gestion commercial/apporteur
    if (GetParamSoc('SO_AfFormatExer') = 'AUC') then result:=result+'71408;'  //duplication N en N+1
       else if (V_PGI.PassWord <> CryptageSt(DayPass(Date)))then  result:=result+'71412;';    //epuration ech facturation
    If not(GetParamSOc('SO_AFGESTIONTARIF'))then
        result:= result+'-71600;'; // pas tarif TTC
    if ((VH_GC.AFFormatExer='AA') or  (VH_GC.AFFormatExer='AM')or (VH_GC.AFFormatExer='A')) then result:=result+'71413;';
    if (GetParamSoc('SO_AfFormatExer') <> 'AUC') then FMenuG.RemoveItem(74109) // pas assistant codif affaire
         else FMenuG.RemoveItem(74602); // pas reprise SCOT pour l'instant
    result:= result +'-74712;';// tablette hirarchique pas ok
    result:= result +'-73601;';// modif boni/mali pas ok
    if Not(V_PGI.SAV) then result:=result+'74111;74651;74656;';
    if Not (V_PGI.PassWord = CryptageSt(DayPass(Date))) then  result:=result+'74108;'; // assistant de création PCL
    if (GetParamsoc('SO_AFLIENDP') =true) then result:=result+'74702;' // groupe utilisateur
    else result:=result+'-74616;74604;74605;74607;74606;'; // fct sur annuaire + autres fct - PL le 18/11/02 ajout du -74616

    {$ifndef DEBUG}
    result:=result+'74701;74703;74704;74705;72706;'; // zone menus utilisateurs et accès
    result := result + '71105;';    //$$$JP - mul tiers avec Annuaire/DPs
    {$endif}

    result:=result+'74457;-74520;74522;74523;';
    if (GetParamSoc('SO_AFAPPAVECBM')<>True) then result:=result+'73605;74207;';
    if (VH_GC.GAAchatSeria = False)  then result:=result+'139200;';   // analyse achats
    end;   // fin spécif SCOT

{$ifndef DEBUG}
   Result:=Result+'74708;74709;73402;73618;74708;74709;';
   if Not(VH_GC.AFGestionCom) then result:=result+'73306;'; // sup commissionnement
{$ENDIF}
{$IFDEF SANSCOMPTA} // fonction ouverture fermeture de compte tiers
   result:=result+'71404;71405;71406;';
{$ENDIF}
{$IFDEF EAGLCLIENT}  // non dispo EAGL
   result:=result+'71403;71410;-71415;71430;71607;71407;';
{$ENDIF}
   if (GetParamSoc('SO_AFAPPCUTOFF')<>True) then  result:=result+'-72700;';
   If (GetParamSoc('SO_AFCUTMODEECLAT') ='GLO') then  result:=result+'72703;';

   if Not(VH_GC.AFGestionCom) then result:=result+'71501;'; // sup apporteurs
   if Not(VH_GC.CleAffaire.GestionAvenant) then result:=result+'-71110;';
   If (( not V_PGI.Superviseur) and ( not V_PGI.Controleur)) or (GetParamsoc('SO_AFTYPECONF')<>'AGR') then
       result:=result+'-74706;74706;74707;'; //  Groupe de confidentialité
   if VH_GC.GAAchatSeria = False then result:=result+'-74560;-74550;';   // table libre fournissuer
   if VH_GC.GAPlanningSeria = False then result:=result+'-74570;-74580;-71260;';   // table libre tâche + edition
   if VH_GC.GASaisieDecSeria = False then result:=result+'74655;74654;';   // export fichier
   if Not(VH_GC.GAPlanningSeria) then result:= result+'-74380;-74570;-74580;';   // Libres tâches
   if (GetParamSoc('SO_AFVISAACTIVITE')<>True) then result:=result+'72501;'; //pas de visa
   if (GetParamSoc('SO_AFAPPCUTOFF')<>True) then result:=result+'138700;'; //pas de cutoff fournisseur
   if (GetParamSoc('SO_AFCUTMODEECLAT')='GLO') then result:=result+'138704;'; //pas cut off detaillé
   result:=result+'72300;';

{try
    TobModelePlanning:=TOB.Create ('les modeles', Nil, -1);
    TobModelePlanning.LoadDetailDB ('HRPARAMPLANNING', '', '',Nil, True) ;
    if (TobModelePlanning.Detail.count < 4) then result:=result+'153204;';  // planning 4
    if (TobModelePlanning.Detail.count < 3) then result:=result+'153203;';   // planning 3
    if (TobModelePlanning.Detail.count < 2) then result:=result+'153202;';   // planning 2
    if (TobModelePlanning.Detail.count < 1) then result:=result+'153201;';  // planning 1
    finally
    TobModelePlanning.Free;
    end;
}                            
  //C.B 14/11/02 vieux plus utilisé
  //if Not(V_PGI.SAV) and not(ctxScot In V_PGI.PGIContexte) then result:=result+'73405;'; // cde => Fac
  If (GetParamSoc('SO_AFFACTPARRES') = 'SAN') then Result:=Result+'73500;'; // pas modif facture éclatée si pas gérée
  if (GetParamSoc('SO_AFGESTIONAPPRECIATION')<>True) then result:=result+'73600;';   // apprec
  if Not(V_PGI.SAV) and not(ctxScot In V_PGI.PGIContexte) then result:=result+'74651;74652;74656;';
  if Not (V_PGI.PassWord = CryptageSt(DayPass(Date))) then result:= Result+'74613;74699;74618;-74610';
  if not(GetParamSoc('SO_PgLienRessource')) then Result :=Result+';71445';  //pa sline ressource/salarié
  result:= result +';72510;72511'; //temporaire faire ensuite test sur line paie
  result := result + ';73408'; // gen avoir globaux pas OK
end;




Procedure TraduireUneTable(Prefixe : string);
Var  NumTable, i : integer;
Mcd : IMCDServiceCOM;
Table     : ITableCOM ;
FieldList : IEnumerator ;
Field     : IFieldCOMEx;
BEGIN
  MCD := TMCD.GetMcd;
	if not mcd.loaded then mcd.WaitLoaded();
 	Table := MCD.GetTable(PrefixeToTable(Prefixe));
 	FieldList := Table.Fields;
	FieldList.Reset();
	While FieldList.MoveNext do
  begin
  	Field := FieldList.Current as IFieldCOMEx ;
    if Field.Name<>'' then
    BEGIN
    Field.Libelle:=TraduitGA(Field.Libelle);
    // traduction des libellés des 3 parties de l'affaire
    if Pos ('AFFAIRE1',Field.Name) > 0 then
    Field.Libelle := VH_GC.CleAffaire.Co1Lib;
    if Pos ('AFFAIRE2',Field.Name) > 0 then
    Field.Libelle := VH_GC.CleAffaire.Co2Lib;
    if Pos ('AFFAIRE3',Field.Name) > 0 then
    Field.Libelle := VH_GC.CleAffaire.Co3Lib;
    END;
  end;
END;

Procedure TraduireMetaDataAffaire;
BEGIN
if not(ctxAffaire in V_PGI.PGIContexte) then Exit;
TraduireUneTable ('AFF');
TraduireUneTable ('ACT'); //mcd 14/10/02 table activite pour cube
 // mettre ci-dessus les tables utilisées dans les cubes,afin que les
 // parties X affaire soient remplacé par les  options des param
 // ne fait pas la traduction GA en GI car la tob utilisé pour ceci est vide
 // A réfléchir ... tout passer par cette fct ??? même si GA ???
If Not(ctxScot in V_PGI.PGIContexte) Then  Exit;
TraduireUneTable ('AFA');
TraduireUneTable ('ARS');
TraduireUneTable ('T');
TraduireUneTable ('ATB');
TraduireUneTable ('GCP'); //mcd 04/05/01
TraduireUneTable ('GL'); //mcd 31/05/02
TraduireUneTable ('GP'); //mcd 31/05/02
TraduireUneTable ('ACU'); //mcd 25/07/02
TraduireUneTable ('APL'); //mcd 26/09/02 planning
TraduireUneTable ('ATA'); //mcd 26/09/02 planning
TraduireUneTable ('ATR'); //mcd 26/09/02 planning
  //mcd 18/07/03 pour traduction paramsoc il faudrait traudire dans
  //paramsoc, le champs soc_design pour  soc_tree like "001;004;%;000;"
  // ... mais écrirait dans la base ??? ext ce vraiment à faire..  Inutile ne plus de
  // le faire à chaque connection. serait à faire uniquement à chaque mise à jour de socref ..
  // en plus le ferait dans notre base de travail donc aussi pour GA ....

end;

end.

