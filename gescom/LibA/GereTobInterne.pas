unit GereTobInterne;

interface
uses  Classes,
{$IFDEF EAGLCLIENT}
//MaineAGL,
{$ELSE}
//fe_main,
{$ENDIF}

AglInit, sysutils, utob, EntGC;

const
  cNbMaxiDeFilles = 50;
  cCaractereMarqueur = '$$';

function MaTobInterne (Identifiant : string) : TOB;
function UneTOBInterne (var Identifiant : string) : TOB;
function DetruitMaTobInterne (Identifiant : string) : boolean;

implementation


{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 07/02/2003
Modifié le ... :   /  /
Description .. : Fonction de création d'une sous TOB personnelle dans
Suite ........ : VH_GC.AFTobInterne
Suite ........ : Entree : Identifiant = identifiant de notre TOB perso
Suite ........ : Sorties : Identifiant = identifiant de notre TOB perso
Suite ........ :            Result = pointeur sur notre TOB perso nouvellement créée ou
Suite ........ :            retrouvée dans VH_GC.AFTobInterne
Suite ........ :
Mots clefs ... : VH_GC;AFTobInterne
*****************************************************************}
function MaTobInterne (Identifiant : string) : TOB;
  begin

  if (VH_GC.AFTobInterne = nil) then
    // On créé VH_GC.AFTobInterne si elle n'existe pas
    begin
      VH_GC.AFTobInterne := TOB.Create('la TOB' + cCaractereMarqueur, nil, -1);
    end;

  // Si VH_GC.AFTobInterne existe, on regarde si on a une fille de nom Identifiant
  if (Not VH_GC.AFTobInterne.TOBNameExist(Identifiant)) then
    begin
      // Si on ne trouve pas de fille, on en créé une et on lui colle un identifiant
      Result :=  TOB.Create (Identifiant, VH_GC.AFTobInterne, -1);
      Result.AddChampSupValeur ('IDENT' + cCaractereMarqueur, cCaractereMarqueur, false);
    end
  else
    begin
      // Si on trouve une fille, on cherche l'identifiant
      Result := VH_GC.AFTobInterne.FindFirst (['IDENT' + cCaractereMarqueur], [cCaractereMarqueur], true);
    end;

  end;

{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 07/02/2003
Modifié le ... :   /  /
Description .. : Fonction de création d'une sous TOB personnelle dans
Suite ........ : VH_GC.AFTobInterne avec un identifiant automatique
Suite ........ : Entree : Identifiant = identifiant de notre TOB perso
Suite ........ : Sorties : Identifiant = identifiant de notre TOB perso
Suite ........ :  s'il était vide en entrée, renvoie un identifiant unique que l'on devra
Suite ........ :  utiliser pour désallouer
Suite ........ :            Result = pointeur sur notre TOB perso
Suite ........ :
Suite ........ :
Mots clefs ... : VH_GC;AFTobInterne
*****************************************************************}
function UneTOBInterne (var Identifiant : string) : TOB;
  var
  IdTempo : string;
  IndiceId : integer;
  begin
  Result := nil;

  // Si Identifiant est vide, on en trouve un unique
  if (Identifiant = '') then
    begin
      IndiceId := 1;
      Identifiant := 'LaTobFille' + cCaractereMarqueur;
      IdTempo := Identifiant + inttostr (IndiceId);
      if (VH_GC.AFTobInterne <> nil) then
        while (VH_GC.AFTobInterne.TOBNameExist (IdTempo)) and (IndiceId <= cNbMaxiDeFilles) do
          begin
            Inc (IndiceId);
            IdTempo := Identifiant + inttostr (IndiceId);
          end;

      // Si on dépasse le nombre maxi de TOB filles de VH_GC.AFTobInterne, on sort à nil
      if (IndiceId > cNbMaxiDeFilles) then exit;
      Identifiant := IdTempo;
    end;

  Result := MaTobInterne (Identifiant);

  end;

{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 07/02/2003
Modifié le ... :   /  /    
Description .. : Fonction de destruction de notre sous TOB personnelle
Suite ........ : Entrée : Identifiant = identifiant de notre TOB perso
Suite ........ : Sortie : Result = il y a eu désallocation d'une TOB perso ou pas (true/false)
Suite ........ :
Suite ........ :
Suite ........ :
Mots clefs ... : VH_GC.AFTobInterne
*****************************************************************}
function DetruitMaTobInterne (Identifiant : string) : boolean;
  var
  TobTempo : TOB;

  begin
  Result := false;

  if (VH_GC.AFTobInterne <> nil) then
    begin
      // Si la fille nommée 'Identifiant' existe, on la free ainsi que toutes ses filles
      if (VH_GC.AFTobInterne.TOBNameExist(Identifiant)) then
        begin
          TobTempo := VH_GC.AFTobInterne.FindFirst (['IDENT' + cCaractereMarqueur], [cCaractereMarqueur], true);
          TobTempo.Free;
          Result := true;
        end;

      // S'il n'y a plus rien dans VH_GC.AFTobInterne, on la free
      if (VH_GC.AFTobInterne.Detail.Count = 0) then
        begin
          VH_GC.AFTobInterne.Free;
          VH_GC.AFTobInterne := nil;
        end;
    end;

  end;

end.
