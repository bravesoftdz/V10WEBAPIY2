unit UtilMenuAff;

// Source permettant de mettre des fct globale GI/GA, sans surcharger AffaireUtil qui est utilisé de partout
// créer 12/06/02 pour passer fct AffdispatchTT de affaireUtil dans ce source
interface
Uses                     
    sysutils, forms,
{$IFDEF EAGLCLIENT}
     MaineAGL,
{$ELSE}
     Fe_Main, hMsgBox,
{$ifndef GIGI}
      Galoutil,ubob, // utilisé par fct spécif GA A laisser
{$ENDIF}
{$ENDIF}
      Hent1,  Hctrls,  AglInit, ConfidentAffaire,UtilArticle,
      UtomAffaire,UtomRessource ,TiersUtil,UtomArticle,CBPPath
     ;
// fct générale pour GA
Procedure AffDispatchTT ( Num : Integer ; Action : TActionFiche ; Lequel,TT, Range : String) ;
Function AfNoAideTablette (TT:string) : integer;
{$ifndef GIGI}
{$IFNDEF EAGLCLIENT}
function IMPORT_BOB(CodeProduit:string): Integer;
{$endif}
{$endif}

Implementation

Procedure AffDispatchTT ( Num : Integer ; Action : TActionFiche ; Lequel,TT, Range : String) ;
var
TypeArticle, Champ, Valeur,Critere :string;
ArticleAff:string;
Arg5LanceFiche:string;
 X,ii, jj : integer;
        
BEGIN
if ((Num=8) and (TT='GCTIERSSAISIE'))  then
    BEGIN
    ii:=TTToNum(TT) ;
    if (ii>0) and (pos('T_NATUREAUXI="FOU"',V_PGI.DECombos[ii].Where)>0) then num:=12  ;
    END ;
     // alimente messgae plus titre ==> peut être changer d'un seul coup
  case Num of
    5 : {affaires}
        BEGIN
        // A- modifié par PL le 23/04/2001 pour homogeinisation des appels
        // les zooms ne fonctionnaient plus...
        // B- modifié le 28/05/2001 pour homogéiniser les appels : on doit passer
        // le statut dans le Range sous la forme 'STATUT=AFF' ou si on ne le passe pas
        // la fonction le détermine par le 1er caractère du code affaire (Lequel)
        Arg5LanceFiche:= ActionToString(Action);
        if (Range ='') then
            begin
            Range:='STATUT:AFF'; //mcd 19/02/01
            if (Lequel<>'') then
                if (Lequel[1]='P') then Range:='STATUT:PRO';
            end;

        Arg5LanceFiche:= Arg5LanceFiche + ';' + Range;

        if AGLJaiLeDroitFiche(['AFFAIRE',ActionToString(Action),Range],3)  then begin
         (* mcd 12/06/02if ctxScot in V_PGI.PGIContexte then
                  AGLLanceFiche('AFF','MISSION','',Lequel,Arg5LanceFiche)
             else
                  AGLLanceFiche('AFF','AFFAIRE','',Lequel,Arg5LanceFiche); *)
           AFLanceFiche_Affaire(lequel,Arg5lancefiche);
           end;
        // fin modifs PL 23/04/2001
        END;
    6 : BEGIN{ressources}
        Arg5LanceFiche:= ActionToString(Action);
        if (Range<>'') then Arg5LanceFiche:= Arg5LanceFiche + ';' + Range;
        if AGLJaiLeDroitFiche(['RESSOURCE',ActionToString(Action)],2)
          then  AFLanceFiche_Ressource(lequel,Arg5lancefiche);
           // mcd 12/06/02 AGLLanceFiche('AFF','RESSOURCE','',Lequel,Arg5LanceFiche);
        END;
    7 : BEGIN {article}
        if not IsCodeArticleUnique(Lequel) then Lequel:=CodeArticleUnique(Lequel,'','','','','') ;
        // Dans le cas de la création, il ne faut pas passer le code article en entrée
           //mcd ajout lequel = '', car si liste vide,on écrase par ligne précédente lequel avec '   X' et on
           // n'arrive pas sur une création de nouvelle fiche sur le typeartciel voulu
        TypeArticle := GetChampsArticle(Lequel, 'GA_TYPEARTICLE'); // mcd 30/04/02 déplacé (était après instruction suivante ..)
        If typeArticle='' then
          begin  ;//mcd 22/09/03 si appel depuis lookup avec condition plus
          jj :=Pos ('GA_TYPEARTICLE',Range);
          If jj <>0 then
            begin
            TypeArticle:= Copy (Range,Pos('"',Copy(Range,jj+15,Strlen(Pchar(Range))))+jj +15,3);
            end;
          end;
        if (Action=taCreat) then begin ArticleAff:=''; lequel:=''; end else ArticleAff:=Lequel;


        Arg5LanceFiche:= ActionToString(Action);
        // Si on n'a ni article, ni type d'article, on suppose que l'on passe l'info manquante dans le range
        if (ArticleAff='') and (TypeArticle='') then
            BEGIN
            Critere:=(ReadTokenSt(Range));
            While (Critere <>'') do
                BEGIN
                X:=pos('=',Critere);
                if x<>0 then
                    begin
                    Champ:=copy(Critere,1,X-1);
                    Valeur:=Copy (Critere,X+1,length(Critere)-X);
                    end;
                if Champ = 'TYPEARTICLE' then TypeArticle := Valeur;
                if (TypeArticle<> '' ) then Arg5LanceFiche:= Arg5LanceFiche + ';TYPEARTICLE='+TypeArticle;
                Critere:=(Trim(ReadTokenSt(Range)));
                END;
 
            if (TypeArticle='') then
               // Type d'article par defaut est PRE si aucun type précisé
               begin
               TypeArticle:='PRE';
               Arg5LanceFiche:= Arg5LanceFiche + ';TYPEARTICLE='+TypeArticle;
               end;
            END
        else
            begin
            Arg5LanceFiche:= Arg5LanceFiche + ';TYPEARTICLE='+TypeArticle;
            if (Range<>'') then Arg5LanceFiche:= Arg5LanceFiche + ';' + Range;
            end;
        if (IsArticleSpecif(TypeArticle)='FicheAffaire') then
           Begin
            if AGLJaiLeDroitFiche(['ARTICLE',ActionToString(Action),typearticle],3) then begin
               if TypeArticle = 'POU' then
                    AFLanceFiche_ArticlePouGA( ArticleAff, Arg5LanceFiche)
               else AFLanceFiche_ArticleGA( ArticleAff, Arg5LanceFiche);
               end;
            End
          else
            AGLLanceFiche('GC', 'GCARTICLE', '', ArticleAff, Arg5LanceFiche) ;
        END ;

    8 : {Clients par T_TIERS}
        begin
        //ATTENTION M: lequel doit contenir T_TIERS
        Arg5LanceFiche:= ActionToString(Action);
        if (Range<>'') then Arg5LanceFiche:= Arg5LanceFiche + ';' + Range;
        if (Arg5LanceFiche <> '') then begin
           X:=pos('T_NATUREAUXI',Arg5LanceFiche);
           if (x=0) then Arg5LanceFiche := Arg5LanceFiche +';T_NATUREAUXI=CLI';
           end
        else Arg5LanceFiche := 'T_NATUREAUXI=CLI';
        if lequel <>'' then    // passe de tiers à auxiliaire
           Lequel := TiersAuxiliaire (Lequel,false);
        if ctxAffaire in V_PGI.PGIContexte then
            BEGIN        //penser de modifier aussi tag 28
            If AglJaiLeDroitFiche (['AFTIERS',ActionToString(Action)],2) then begin
              AGLLanceFiche('GC','GCTIERS','',Lequel,Arg5LanceFiche);
             end ;
            END
        else
            AGLLanceFiche('GC','GCTIERS','',Lequel,ActionToString(Action)+';MONOFICHE;T_NATUREAUXI=CLI') ;
        end ;
   12 : {Fournisseurs via T_TIERS} begin
            Lequel:=TiersAuxiliaire(Lequel,false);
            AGLLanceFiche('GC','GCFOURNISSEUR','',Lequel,ActionToString(Action)+';MONOFICHE;T_NATUREAUXI=FOU') ;
            end;
   16 : {contacts}
        BEGIN
        Arg5LanceFiche:= ActionToString(Action);
        if (Range<>'') then Arg5LanceFiche:= Arg5LanceFiche + ';' + Range;
//        if (TT<>'') then Arg5LanceFiche:= Arg5LanceFiche + ';' + TT;
        if  AglJaiLeDroitFiche (['CONTACT',ActionToString(Action)],2)
            then begin
            AGLLanceFiche('YY','YYCONTACT', Range,'',Arg5LanceFiche) ;
            end;
       end;
   28 : {Clients par T_AUXILIAIRE}
        BEGIN
        Arg5LanceFiche:= ActionToString(Action);
        if (Range<>'') then Arg5LanceFiche:= Arg5LanceFiche + ';' + Range;
        if (Arg5LanceFiche <> '') then begin
           X:=pos('T_NATUREAUXI',Arg5LanceFiche);
           if (x=0) then Arg5LanceFiche := Arg5LanceFiche +';T_NATUREAUXI=CLI';
           end
        else Arg5LanceFiche := 'T_NATUREAUXI=CLI';
        if ctxAffaire in V_PGI.PGIContexte then
            BEGIN    //penser de modifier aussi tag 8
            If AglJaiLeDroitFiche (['AFTIERS',ActionToString(Action)],2) then begin
              AGLLanceFiche('GC','GCTIERS','',Lequel,Arg5LanceFiche);
              end ;
            END
        else
            AGLLanceFiche('GC','GCTIERS','',Lequel,ActionToString(Action)+';MONOFICHE;T_NATUREAUXI=CLI') ;
         END ;
 end ;
END ; 

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/12/2002
Modifié le ... :   /  /
Description .. :  IMPORTATION DES BOBS
//-----------------------------------------------------------------------
// LE NOM DES BOB SE COMPOSE DE
// - Code Produit   XXXX
// - Num version base 9999
// - type de BOB (F:fiche,M:Menu,D:data);
// - Num version 999
// - extension .BOB
//-----------------------------------------------------------------------
//fichier actuel CGAS50595D001.BOB
//-----------------------------------------------------------------------
Mots clefs ... :
*****************************************************************}
{$ifndef GIGI}
{$ifndef EAGLCLIENT}
function IMPORT_BOB(CodeProduit:string): Integer;
var sFileBOB   : string;
    Chemin     : string;
    SearchRec  : TSearchRec;
    ret        : integer;

BEGIN
  Result := 0;
//  Chemin := ExtractFileDrive(Application.ExeName) + '\PGI00\BOB\'+ CodeProduit +'\';
  Chemin := IncludeTrailingBackSlash(ExtractFileDrive(Application.ExeName)) + IncludeTrailingBackSlash(ExtractFilePath (TCBPPath.GetCegidDistriBob))+ CodeProduit +'\';
  ret := FindFirst(Chemin+CodeProduit+'*.BOB', faAnyFile, SearchRec);
  while ret = 0 do
  begin
    //RECUPERE NOM DU BOB
    sFileBOB := SearchRec.Name;
//    case TestAGLIntegreBob(Chemin + sFileBOB) of
    case AGLIntegreBob(Chemin + sFileBOB,FALSE,TRUE) of
       0  :// OK
           begin
             if V_PGI.SAV then Pgiinfo('Intégration de : '+sFileBOB, TitreHalley);//Resultif not LIA_JOURNAL_EVENEMENT(sTempo) then Result := -1;
             if copy(sFileBob,10,1) = 'M' then Result := 1; //SI BOB AVEC MENU, ON REND 1 POUR SORTIR DE L'APPLICATION
           end;

      1  : if V_PGI.SAV then Pgiinfo('Intégration déjà effectuée :'+sFileBOB, TitreHalley);// Intégration déjà effectuée

      // Erreur d'écriture dans la table YMYBOBS
      -1  : if V_PGI.SAV then PGIInfo('Erreur d''écriture dans la table YMYBOBS :'+Chemin + sFileBOB,'IMPORT_BOB');

      // Erreur d'intégration dans la fonction AglImportBob
      -2  : if V_PGI.SAV then PGIInfo('Erreur d''intégration dans la fonction AglImportBob :'+Chemin + sFileBOB,'PCL_IMPORT_BOB');

      //Erreur de lecture du fichier BOB.
      -3  : if V_PGI.SAV then PGIInfo('Erreur de lecture du fichier BOB :'+Chemin + sFileBOB,'IMPORT_BOB');

      // Erreur inconnue.
      -4  : if V_PGI.SAV then PGIInfo('Erreur inconnue :'+Chemin + sFileBOB,'IMPORT_BOB');
    end;
    ret := FindNext(SearchRec);
  end;
  sysutils.FindClose(SearchRec);
END;
{$endif}
{$endif}


{***********A.G.L.***********************************************
Auteur  ...... : MC DESSEIGNET
Créé le ...... : 15/05/2003
Modifié le ... :   /  /    
Description .. : Focntion qui rend le n° aide correspodnant à la tablette
Mots clefs ... : AIDE;TABLETTE
*****************************************************************}
Function AfNoAideTablette (TT:string) : integer;
begin
  result :=0;
  If TT = 'TTREGIMETVA' then Result :=  110000183
  else if TT = 'GCLIBRETIERS1' then Result := 110000252
  else if TT = 'GCLIBRETIERS2' then Result := 110000252
  else if TT = 'GCLIBRETIERS3' then Result := 110000252
  else if TT = 'GCLIBRETIERS4' then Result := 110000252
  else if TT = 'GCLIBRETIERS5' then Result := 110000252
  else if TT = 'GCLIBRETIERS6' then Result := 110000252
  else if TT = 'GCLIBRETIERS7' then Result := 110000252
  else if TT = 'GCLIBRETIERS8' then Result := 110000252
  else if TT = 'GCLIBRETIERS9' then Result := 110000252
  else if TT = 'GCLIBRETIERSA' then Result := 110000252
  else if TT = 'TTCIVILITE' then Result := 110000189
  else if TT = 'TTLANGUE' then Result := 110000190
  else if TT = 'GCEMPLOIBLOB' then Result := 110000172
  else if TT = 'GCCOMMENTAIREENT' then Result := 110000167
  else if TT = 'GCCOMMENTAIREPIED' then Result := 110000167
  else if TT = 'GCCOMMENTAIRELIGNE' then Result := 110000167
  else if TT = 'GCCOMPTATIERS' then Result := 110000156
  else if TT = 'GCZONECOM' then Result := 110000156
  else if TT = 'TTSECTEUR' then Result := 110000156
  else if TT = 'GCORIGINETIERS' then Result := 110000156
  else if TT = 'TTFONCTION' then Result := 1125000
  else if TT = 'TTLIENPARENT' then Result := 1125000
  else if TT = 'YYSERVICE' then Result := 1125000
  else if TT = 'TTTARIFCLIENT' then Result := 110000156
  else if TT = 'GCFAMILLENIV1' then Result := 110000182
  else if TT = 'GCFAMILLENIV2' then Result := 110000182
  else if TT = 'GCFAMILLENIV3' then Result := 110000182
  else if TT = 'GCLIBFAMILLE' then Result := 110000182
  else if TT = 'GCCOMPTAARTICLE' then Result := 110000182
  else if TT = 'GCTARIFARTICLE' then Result := 110000182
  else if TT = 'GCLIBREART1' then result := 110000251
  else if TT = 'GCLIBREART2' then result := 110000251
  else if TT = 'GCLIBREART3' then result := 110000251
  else if TT = 'GCLIBREART4' then result := 110000251
  else if TT = 'GCLIBREART5' then result := 110000251
  else if TT = 'GCLIBREART6' then result := 110000251
  else if TT = 'GCLIBREART7' then result := 110000251
  else if TT = 'GCLIBREART8' then result := 110000251
  else if TT = 'GCLIBREART9' then result := 110000251
  else if TT = 'GCLIBREARTA' then result := 110000251
  else if TT = 'GCLIBREPIECE1' then Result := 110000256
  else if TT = 'GCLIBREPIECE2' then Result := 110000256
  else if TT = 'GCLIBREPIECE3' then Result := 110000256
  else if TT = 'GCLIBREFOU1' then Result := 110000253
  else if TT = 'GCLIBREFOU2' then Result := 110000253
  else if TT = 'GCLIBREFOU3' then Result := 110000253
  else if TT = 'YYLIBREET1' then Result := 110000255
  else if TT = 'YYLIBREET2' then Result := 110000255
  else if TT = 'YYLIBREET3' then Result := 110000255
  else if TT = 'YYLIBREET4' then Result := 110000255
  else if TT = 'YYLIBREET5' then Result := 110000255
  else if TT = 'YYLIBREET6' then Result := 110000255
  else if TT = 'YYLIBREET7' then Result := 110000255
  else if TT = 'YYLIBREET8' then Result := 110000255
  else if TT = 'YYLIBREET9' then Result := 110000255
  else if TT = 'YYLIBREETA' then Result := 110000255
  else if TT = 'YYLIBRECON1' then Result := 110000254
  else if TT = 'YYLIBRECON2' then Result := 110000254
  else if TT = 'YYLIBRECON3' then Result := 110000254
  else if TT = 'YYLIBRECON4' then Result := 110000254
  else if TT = 'YYLIBRECON5' then Result := 110000254
  else if TT = 'YYLIBRECON6' then Result := 110000254
  else if TT = 'YYLIBRECON7' then Result := 110000254
  else if TT = 'YYLIBRECON8' then Result := 110000254
  else if TT = 'YYLIBRECON9' then Result := 110000254
  else if TT = 'YYLIBRECONA' then Result := 110000254
  else if TT = 'AFJUSTIFBONI' then Result := 120000175
  else if TT = 'AFCOMPTAAFFAIRE' then Result := 120000157
  else if TT = 'AFTLIENAFFTIERS' then Result := 120000160
  else if TT = 'AFDEPARTEMENT' then Result := 120000159
  else if TT = 'AFTRESILAFF' then Result := 120000161
  else if TT = 'AFFAIREPART1' then Result := 120000156
  else if TT = 'AFFAIREPART2' then Result := 120000156
  else if TT = 'AFFAIREPART3' then Result := 120000156
  else if TT = 'AFETATAFFAIRE' then Result := 120000158
  else if TT = 'AFTREGROUPEFACT' then Result := 120000162
  else if TT = 'AFTNIVEAUDIPLOME' then Result := 120000166
  else if TT = 'AFTTYPERESSOURCE' then Result := 120000163
  else if TT = 'AFTSTANDCALEN' then Result := 120000169
  else if TT = 'AFTTYPEHEURE' then Result := 120000167
  else if TT = 'AFTLIBREAFF1' then Result := 120000153
  else if TT = 'AFTLIBREAFF2' then Result := 120000153
  else if TT = 'AFTLIBREAFF3' then Result := 120000153
  else if TT = 'AFTLIBREAFF4' then Result := 120000153
  else if TT = 'AFTLIBREAFF5' then Result := 120000153
  else if TT = 'AFTLIBREAFF6' then Result := 120000153
  else if TT = 'AFTLIBREAFF7' then Result := 120000153
  else if TT = 'AFTLIBREAFF8' then Result := 120000153
  else if TT = 'AFTLIBREAFF9' then Result := 120000153
  else if TT = 'AFTLIBREAFFA' then Result := 120000153
  else if TT = 'AFTLIBRERES1' then Result := 120000155
  else if TT = 'AFTLIBRERES2' then Result := 120000155
  else if TT = 'AFTLIBRERES3' then Result := 120000155
  else if TT = 'AFTLIBRERES4' then Result := 120000155
  else if TT = 'AFTLIBRERES5' then Result := 120000155
  else if TT = 'AFTLIBRERES6' then Result := 120000155
  else if TT = 'AFTLIBRERES7' then Result := 120000155
  else if TT = 'AFTLIBRERES8' then Result := 120000155
  else if TT = 'AFTLIBRERES9' then Result := 120000155
  else if TT = 'AFTLIBRERESA' then Result := 120000155
  else if TT = 'AFPROFILGENER' then Result := 120000221
  else if TT = 'TTFORMEJURIDIQUE' then Result := 110000187;
end;

end.

