import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Currency;
import com.jacob.com.LibraryLoader;
import com.jacob.com.Variant;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;


public class KrossImport {

    private List<String> allLines = new ArrayList<>();
    private List<String> productSymbols = new ArrayList<>();
    private List<String> productQuantities = new ArrayList<>();
    private List<String> nettoPrices = new ArrayList<>();
    private String invoiceNumber = "";

    public List<String> readAllLinesFromFile() throws IOException {
        allLines = Files.readAllLines(Paths.get("c:\\ImportFaktur\\kross.txt"));
        return allLines;
    }

    public String returnInvoiceNumber() {
        invoiceNumber = allLines.stream()
                .filter(x -> x.contains("Numer faktury:"))
                .map(x -> x.replace("Numer faktury: ", ""))
                .collect(Collectors.joining());
        return invoiceNumber;
    }

    public List<String> returnProductSymbols() {
        Pattern p = Pattern.compile("((?:[A-Z]+[0-9]+|[0-9]+[A-Z]+)[A-Z0-9]+)(\\W+[0-9]+)(\\W+)([K][P][L]|[S][Z][T]|[P][A][R])");
        Matcher m;
        for (int i = 0; i < allLines.size(); i++) {
            m = p.matcher(allLines.get(i));
            if (m.find()) {
                productSymbols.add(m.group(1));
            }
        }
        return productSymbols;
    }

    public List<String> returnProductQuantities() {
        Pattern p = Pattern.compile("([0-9]+)([\\W])+([K][P][L]|[S][Z][T]|[P][A][R])");
        Matcher m;
        for (int i = 0; i < allLines.size(); i++) {
            m = p.matcher(allLines.get(i));
            if (m.find()) {
                productQuantities.add(m.group(1));
            }
        }
        return productQuantities;
    }

    public List<String> returnNettoPrices() {
        Pattern p = Pattern.compile("([2][3]\\W*)([0-9]+\\W*[,][0-9]+|[0-9]+\\s[0-9]+[,][0-9]+)");
        Matcher m;
        for (int i = 0; i < allLines.size(); i++) {
            m = p.matcher(allLines.get(i));
            if (m.find()) {
                String s = m.group(2).replaceAll("\\s+", "");
                nettoPrices.add(s);
            }
        }
        return nettoPrices;
    }


    public void readDataFromInvoice() throws IOException {
        readAllLinesFromFile();
        returnInvoiceNumber();
        returnProductSymbols();
        returnProductQuantities();
        returnNettoPrices();
    }


    public static void main(String[] args) throws IOException {
        //utworzenie streamu zapisującego system.out.println do pliku log.txt
        File file = new File("C:\\ImportFaktur\\log_kross.txt");
        PrintStream stream = new PrintStream(file);
        System.setOut(stream);


        //utworzenie obiektu faktury i zaczytanie danych z faktury
        KrossImport faktura = new KrossImport();
        faktura.readDataFromInvoice();


        // zainicjowanie i uruchomienie mostka JACOB
        String libFile = System.getProperty("os.arch").equals("amd64") ? "jacob-1.19-x64.dll" : "jacob-1.19-x86.dll";
        try {
            InputStream inputStream = JacobExample.class.getResourceAsStream(libFile);
            File temporaryDll = new File("c://ImportFaktur//jacob.dll");
            FileOutputStream outputStream = new FileOutputStream(temporaryDll);
            byte[] array = new byte[8192];
            for (int i = inputStream.read(array); i != -1; i = inputStream.read(array)) {
                outputStream.write(array, 0, i);
            }
            outputStream.close();
            System.setProperty(LibraryLoader.JACOB_DLL_PATH, temporaryDll.getAbsolutePath());
            LibraryLoader.loadJacobLibrary();

            //zainicjowanie obiektu subiekta
            ActiveXComponent oSubiekt;
            ActiveXComponent oGT;
            ComThread.InitSTA();
            oGT = new ActiveXComponent("InsERT.GT");

            //podłączenie do bazy danych subiekta
            oGT.setProperty("Konfiguracja", "C:\\ProgramData\\InsERT\\InsERT GT\\Subiekt.xml");

            //uruchomienie subiekta gt
            oSubiekt = oGT.invokeGetComponent("Uruchom", new Variant(0), new Variant(0));


            // wprowadzenie faktury zakupu
            //wywołanie menedżera dokumentów
            ActiveXComponent oFZs = oSubiekt.invokeGetComponent("SuDokumentyManager");
            //wywołanie funkcji "dodaj fakturę zakupu"
            ActiveXComponent oFZ = oFZs.invokeGetComponent("DodajFZ");
            //wklejenie numeru faktury
            oFZ.setProperty("NumerOryginalny", faktura.invoiceNumber);
            //wybór dostawcy na fakturze
            oFZ.setProperty("KontrahentId", "761-14-02-748");
            //utworzenie obiektu Towary
            ActiveXComponent Towary = oSubiekt.invokeGetComponent("Towary");

            //pętla wprowadzająca kolejno pozycje na fakturę
            System.out.println(faktura.productSymbols.size());
            for (int i = 0; i < faktura.productSymbols.size(); i++) {
                //sprawdzenie czy dana pozycja istnieje w bazie towarów
                Variant istnieje = Towary.invoke("Istnieje", faktura.productSymbols.get(i));
                //jeśli pozycja istnieje to wykonaj : ...
                if (istnieje.getBoolean()) {
                    //zaczytanie ceny z ostatniej dostawy produktu o numerze "i"
                    Currency currency = Towary.invokeGetComponent("Wczytaj", new Variant(faktura.productSymbols.get(i))).invokeGetComponent("Zakupy").invokeGetComponent("Wczytaj", new Variant(1)).invoke("Wartosc").getCurrency();
                    double cenaZOstatniejDostawy = (double) currency.longValue() / 10000;
                    //zamiana przecinka na kropkę w wartości netto produktu o numerze "i"
                    String s = faktura.nettoPrices.get(i).replaceAll("[,]", ".");
                    //wyliczenie ceny netto, cena netto = wartość netto dzielona przez ilość i zaokrąglenie ceny do 2 miejsc po przecinku
                    Double f = Double.parseDouble(s) / Double.parseDouble(faktura.productQuantities.get(i));
                    DecimalFormat df = new DecimalFormat("###.##");
                    Double cenaZTejFaktury = Double.parseDouble(df.format(f).replaceAll("[,]", "."));
                    //jeżeli cena produktu z aktualnej faktury jest większa od ceny produktu z ostatniej dostawy to wyświetl komunikat
                    if (cenaZOstatniejDostawy < cenaZTejFaktury) {
                        System.out.println(faktura.productSymbols.get(i) + "  zdrożał o " + (cenaZTejFaktury - cenaZOstatniejDostawy) + " zł");
                    }
                    //dodawanie do faktury zakupu pozycji o określonym kodzie, ustawianie ilości sztuk oraz wartość netto produktów
                    ActiveXComponent danaPozycja = oFZ.invokeGetComponent("Pozycje").invokeGetComponent("Dodaj", new Variant(faktura.productSymbols.get(i)));
                    danaPozycja.setProperty("IloscJm", faktura.productQuantities.get(i));
                    danaPozycja.setProperty("WartoscNettoPoRabacie", faktura.nettoPrices.get(i));
                    //jeśli pozycja nie istnieje to wykonaj : ...
                } else {
                    System.out.println("brak towaru w bazie, dodawanie towaru...");
                    //wywołanie metody dodania nowego towaru do bazy towarów
                    ActiveXComponent Towar = Towary.invokeGetComponent("Dodaj", new Variant(1));
                    //ustawienie nazwy towaru na "NEW", symbolu towaru , i zapisanie towaru w bazie
                    Towar.setProperty("Nazwa", "NEW");
                    Towar.setProperty("Symbol", faktura.productSymbols.get(i));
                    Towar.invoke("Zapisz");
                    //dodanie do faktury pozycji z uprzednio wprowadzonym do bazy towarem
                    ActiveXComponent danaPozycja = oFZ.invokeGetComponent("Pozycje").invokeGetComponent("Dodaj", new Variant(faktura.productSymbols.get(i)));
                    //ustawianie ilości sztuk oraz wartość netto produktu
                    danaPozycja.setProperty("IloscJm", faktura.productQuantities.get(i));
                    danaPozycja.setProperty("WartoscNettoPoRabacie", faktura.nettoPrices.get(i));
                }
            }

            //wyświetlenie tak sporządzonej faktury, bez zapisywania
            oFZ.invoke("Wyswietl");

            //kasowanie tymczasowego pliku biblioteki
            System.out.println("Plik jacob.dll skasowany ? : " + temporaryDll.delete());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}

