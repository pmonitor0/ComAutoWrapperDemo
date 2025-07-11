# ComAutoWrapperDemo

Ez a projekt egy **WPF alapú demó** a [ComAutoWrapper](https://www.nuget.org/packages/ComAutoWrapper) NuGet csomag használatához.  
Megmutatja, hogyan lehet COM objektumokat (pl. Excel) használni Interop DLL-ek nélkül, egyszerű, típusos és visszafogott módon.

## Főbb funkciók

- Excel automatizálás WPF-ből
- COM tulajdonságok (`GetProperty<T>`, `SetProperty`)
- COM metódusok hívása (`CallMethod<T>`)
- COM tagok introspektálása:  
  - Metódusok (`Method`)
  - Olvasható (`PropertyGet`)
  - Írható (`PropertySet`) property-k

## A demó működése

Indítás után a WPF alkalmazás:

1. Console módon automatikusan:
   - Elindítja az Excelt
   - Kitölti szorzótáblával a cellákat
   - Lekérdezi a `Worksheet` COM tagjait
   - Kiírja őket a konzolra
2. Ezután megjeleníti a `MainWindow`-t, vagy kilép, ha nincs további felhasználói interakció

## Telepítés

1. Klónozd a repót:
```bash
git clone https://github.com/pmonitor0/ComAutoWrapperDemo.git
2. Nyisd meg Visual Studio-ban (.sln fájl).

3. Ellenőrizd, hogy a NuGet csomag (ComAutoWrapper) telepítve van.

## Full Excel + Word automation demo

This WPF app runs both Excel and Word COM automation examples without any Interop DLLs:

- Writes data into Excel
- Formats Word paragraph
- Inspects COM members via `ComTypeInspector`

Source: [ComAutoWrapperDemo](https://github.com/pmonitor0/ComAutoWrapperDemo)

Követelmények
Windows (COM miatt)

.NET 6/7/8/9

Telepített Microsoft Excel
Telepített Microsoft Word

Kapcsolódó projekt
[ComAutoWrapper (NuGet)](https://www.nuget.org/packages/ComAutoWrapper)
[ComAutoWrapper (GitHub)](https://github.com/pmonitor0/ComAutoWrapper)

Köszönetnyilvánítás
A projekt ötlete közös fejlesztés eredménye, a ChatGPT támogatásával.

License
MIT