# Librus AutoMail

Skrypt cyklicznie sprawdza pierwszą stronę odebranych wiadomości w Librusie i przesyła nowe, nieprzeczytane wiadomości na wskazane adresy e-mail przez Outlook SMTP z OAuth2.

## Jak działa

- obsługuje wiele kont Librus;
- zachowuje sesje Librus na dysku i loguje ich wiek;
- po wygaśnięciu sesji loguje się ponownie loginem i hasłem;
- zapisuje stan wiadomości w SQLite i oznacza wiadomość jako wysłaną dopiero po przyjęciu jej przez SMTP;
- po obsłużeniu wszystkich kont jeden raz ponawia konta zakończone błędem;
- jest przygotowany do uruchamiania przez timer systemd co 30 minut.

## Wymagania

- Python 3.10 lub nowszy;
- konto Microsoft/Outlook z publiczną aplikacją Entra skonfigurowaną do OAuth2 SMTP;
- konto lub konta Librus Synergia.

## Instalacja

Wszystkie polecenia wykonuj jako zwykły użytkownik Linuksa. Projekt zostanie umieszczony w katalogu domowym aktualnie zalogowanego użytkownika, więc nie trzeba wpisywać ani zmieniać nazwy jego konta.

```bash
cd "$HOME"
git clone https://github.com/sylwesterdec/librus-automail.git
cd "$HOME/librus-automail"
python3 -m venv .venv
./.venv/bin/python -m pip install --upgrade pip
./.venv/bin/python -m pip install -r requirements.txt
cp librus_konta.example.json librus_konta.json
cp .env.example librus-automail.env
chmod 600 librus_konta.json librus-automail.env
```

Następnie edytuj dwa utworzone pliki:

1. W `librus-automail.env` wpisz identyfikator aplikacji Microsoft (`SMTP_CLIENT_ID`), adres konta wysyłającego (`SMTP_EMAIL`) oraz hasło Librus (`LIBRUS_PASSWORD_1`).
2. W `librus_konta.json` wpisz login Librus, nazwę konta widoczną w temacie wiadomości i adresy odbiorców.

Pole `"password": "$LIBRUS_PASSWORD_1"` pozostaw bez zmian — aplikacja pobierze właściwe hasło z prywatnego pliku `librus-automail.env`. Aby dodać kolejne konto, skopiuj obiekt wewnątrz listy `accounts`, użyj kolejnej zmiennej, na przykład `$LIBRUS_PASSWORD_2`, i dodaj ją także do pliku środowiskowego.

Prawdziwa konfiguracja, hasła, tokeny, baza stanu i logi są wykluczone przez `.gitignore` i nie zostaną przypadkowo zapisane w repozytorium.

## Pierwsza autoryzacja Outlook

Pierwsze uruchomienie wymaga interaktywnego logowania Microsoft na komputerze z przeglądarką:

```bash
set -a
source ./librus-automail.env
set +a
./.venv/bin/python ./librus_automail.py --bootstrap-smtp
```

Po udanym logowaniu aplikacja zapisze lokalną pamięć tokenu SMTP. Następnie sprawdź dostęp do Outlooka i wszystkich kont Librus bez czytania treści oraz bez wysyłania poczty:

```bash
./.venv/bin/python ./librus_automail.py --check
```

Jeżeli kontrola zakończy się powodzeniem, wykonaj zwykłe jednorazowe uruchomienie:

```bash
./.venv/bin/python ./librus_automail.py
```

## Timer systemd

Jest to usługa użytkownika systemd. Zapis `%h` w pliku usługi systemd sam wskazuje katalog domowy aktualnego użytkownika. Nie edytuj nazwy użytkownika ani ścieżek.

```bash
mkdir -p "$HOME/.config/systemd/user"
cp systemd/librus-automail.service systemd/librus-automail.timer "$HOME/.config/systemd/user/"
systemctl --user daemon-reload
systemctl --user enable --now librus-automail.timer
systemctl --user list-timers librus-automail.timer
```

Timer uruchamia aplikację o pełnej i wpół do każdej godziny. Aby działał także po wylogowaniu użytkownika, włącz pozostawianie jego usług w tle:

```bash
sudo loginctl enable-linger "$USER"
```

Ręczny test usługi i podgląd logów:

```bash
systemctl --user start librus-automail.service
systemctl --user status librus-automail.service --no-pager
journalctl --user -u librus-automail.service -n 100 --no-pager
tail -n 100 "$HOME/librus-automail/librus_email.log"
```

## Testy

```bash
./.venv/bin/python -m unittest -v test_librus_automail.py
```


## Informacja o wykorzystaniu AI

Projekt powstał przy wsparciu narzędzia sztucznej inteligencji OpenAI Codex. AI pomogła w analizie wymagań, przygotowaniu kodu, testów oraz dokumentacji. Przed publikacją rezultat został sprawdzony i przetestowany przez opiekuna projektu. Informację zamieszczono w celu przejrzystego ujawnienia udziału AI, również w kontekście zasad przejrzystości unijnego AI Act.
