import 'package:gsheets/gsheets.dart';
import 'package:intl/intl.dart'; // Importing the intl package for date formatting

class GoogleSheetsService {
  // Credentials
  static const _credentials = {
    "type": "service_account",
    "project_id": "icftes-app",
    "private_key_id": "7b4a51fc3a197ec3b7015fe8dbb812e60998903c",
    "private_key": "-----BEGIN PRIVATE KEY-----\n"
        "MIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQCitc0fL+I4SMkO\n"
        "UUEK8i7bnvjAtrac7EC3G6Yl6l4smcqYqaeZfF3Ynar7K/GEStA9TkUbL5WikTm1\n"
        "5j39yhy4+yl8lTGp1pgfx2Rdm41YbFgHsg7UFaisPwX57H/Srbu/m7Ekw5eN6W1k\n"
        "iyYt7dmQxpD49ybAikVpp+uxO0o7s5m1V9c+6woV/5lugsPFEq318/f4f1plK3yu\n"
        "PptNVB0vi9ZXTCI+jjU8foTXUOZ6JpZl9N/5EzRxwGg24ivbdhfEkWPaAhzAbliv\n"
        "SOgMQiEwKzZpESb9dH219ISx5Desn6d2UKzDsBpkgailtSazlho/GcHj66CDb6nV\n"
        "8LjV/2wDAgMBAAECggEAIPMO9FNWjM5UhMU4ljZf/dKODjyR82o2Wr5LIZ957a9B\n"
        "IzQr//165axcHRwTfxZFYDzS6sPymeat2KOlBxlgQqd+CcAOvBV8XecbcIdZEsBx\n"
        "/TD2JsWyEBt9ItTdN7U98XneYBMJxE+yeutg0mk5p0NGxVwLaW82ykQaixv2Fuuf\n"
        "ghe4POrlXsIBXqsNBFWONyafWAycbyd665rT8Bei1NjioTvjBVBH5rv6Rq10htaJ\n"
        "UZxezKIvd6KXZjXRTjHrpFQfet3lq3mY8qMUrNkq4r3LH+BEzfMBnXfdKUZWKmWf\n"
        "IRhcWcZBD+9zeF2tB2E+24POx0lSMDvU0Tf1ZzaonQKBgQDbhoN9CMa7qlddMvue\n"
        "ZkeqTIpnKtR12eH6xu3YRSWMHvkRUa8fqHkK3Wt9zJtyu6dLejweFrLSzs54W6yq\n"
        "kDdVapgIR258xRKnYd+z3evUkrqUDne+e0CsJix2yPc3w1LWg9pA+b/e7T0pew85\n"
        "Hsmp065zdA2dA7l9SZ2kuqKZnQKBgQC9vqgsVDE0lsf5Px4GRySLnCAD2h6oorTI\n"
        "x/IvUJpMu29kqL5xhpPpkXN5filPb4tV/XbgPu12Z7JFfGqbzpobiA0r8nUPYYUF\n"
        "kFhFdMGclMkC7Fyw8il85B2IdnZSivgWNhxjePMZlI7vwz6E8cwitP7f7QZjEUKD\n"
        "cjh3ZWJ6HwKBgD5UgyENTOAcDZI415iyEccY1HNWhdywcKlzsjSl7XNLmAyC1OZ4\n"
        "P2YGWG7vmXOKNIYJvugMKdoRPi6OWQhUymFGUsSHA6gJjLJZ59p6OGuy/absNLOw\n"
        "6zv12sofZZI/s1WVOnMYdpIlaihM+JWPWFMP94hwey0J0bDxJgGPvHtBAoGBAIgk\n"
        "d2gvFIsmMN+uoO1iOF+Pqwz4gQ0AiXSSujumur+ZsShpRxQPuqtY+KDQm/VqFHCj\n"
        "h5sIq7tMVgYzag7XI43jhYfl1IYvs5E1a5CSYKTnwH6/dxZi+s+ooWQbk3RQUAcn\n"
        "1iCtVMgi5pgz3/TlxVGVylaDLBUC+lV0K/3HGeyDAoGAViwjiG55VW2cMWgHulmv\n"
        "gxaXMraxZk4vEVApn406xZTc/yBz3o/SLzqu+UzpL8vHSmOQXejg4Iq9Qf/P+HoY\n"
        "Q4NftRCnv5JHN1+3dH9lcBclLaME2olzTLVqtCeSeCZHcYJHs/8vkOKqiboo+pTL\n"
        "wk6vDLB2GAqb3fYWZNF5KWk=\n-----END PRIVATE KEY-----\n",
    "client_email": "icftes@icftes-app.iam.gserviceaccount.com",
    "client_id": "112288024848976293452",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url":
        "https://www.googleapis.com/oauth2/v1/certs",
    "client_x509_cert_url":
        "https://www.googleapis.com/robot/v1/metadata/x509/icftes%40icftes-app.iam.gserviceaccount.com",
    "universe_domain": "googleapis.com"
  };

  // Your spreadsheet ID
  static const _spreadsheetId = '1qglftjbYoytykTbP-FXiGpF1aHR9dSiVJiVn4UeNY0M';

  //####################################################### GET ALL GUEST ##############################

  static Future<Map<String, String>> getGuestNamesWithIDs() async {
    try {
      // Initialize GSheets
      final gsheets = GSheets(_credentials);

      // Fetch spreadsheet by its ID
      final ss = await gsheets.spreadsheet(_spreadsheetId);

      // Get the worksheet by its title
      final sheet = ss.worksheetByTitle('guest_list');

      // Ensure the sheet is not null
      if (sheet != null) {
        // Fetch names and IDs from the sheet (2nd and 1st columns) and sort them alphabetically
        final names = await sheet.values.columnByKey('Name') ?? [];
        final ids = await sheet.values.columnByKey('ID') ?? [];
        final Map<String, String> guestMap = {};

        // Combine names and IDs into a map
        for (int i = 0; i < names.length; i++) {
          if (names[i].isNotEmpty && ids[i].isNotEmpty) {
            guestMap[names[i]] = ids[i];
          }
        }

        return guestMap;
      } else {
        throw ('Worksheet not found');
      }
    } catch (e) {
      throw ('Error fetching guest names: $e');
    }
  }

  //####################################################### CHECK IF GUEST EXIST #######################

  static Future<Map<String, dynamic>> checkIfGuestExists_reg(
      String id, String sheetName) async {
    try {
      final gsheets = GSheets(_credentials);
      final ss = await gsheets.spreadsheet(_spreadsheetId);
      final sheet = ss.worksheetByTitle(sheetName);
      if (sheet != null) {
        final idColumn = await sheet.values.columnByKey('ID') ?? [];
        if (idColumn.contains(id)) {
          final index = idColumn.indexOf(id);
          final names = await sheet.values.columnByKey('Name') ?? [];
          if (index >= 0 && index < names.length && names[index].isNotEmpty) {
            final name = names[index];
            return {'if_exists': true, 'name': name, 'index': index};
          }
        } else if (RegExp(r'^ICFTES_\d{4}$').hasMatch(id)) {
          // Check if the ID is valid based on the format
          return {'if_exists': false, 'name': 'none', 'index': -1};
        }
      }
    } catch (e) {
      throw ('Error checking if guest exists: $e');
    }
    return {'if_exists': null, 'name': 'none', 'index': 0};
  }

  static Future<Map<String, dynamic>> checkIfGuestExists_update(
      String id, String sheetName) async {
    try {
      final gsheets = GSheets(_credentials);
      final ss = await gsheets.spreadsheet(_spreadsheetId);
      final sheet = ss.worksheetByTitle(sheetName);
      if (sheet != null) {
        final idColumn = await sheet.values.columnByKey('ID') ?? [];
        if (idColumn.contains(id)) {
          final index = idColumn.indexOf(id);
          final names = await sheet.values.columnByKey('Name') ?? [];
          if (index >= 0 && index < names.length && names[index].isNotEmpty) {
            final name = names[index];
            return {'if_exists': true, 'name': name, 'index': index};
          }
        }
      }
    } catch (e) {
      throw ('Error checking if guest exists: $e');
    }
    return {'if_exists': false, 'name': 'none', 'index': 0};
  }


  //####################################################### ADD NEW GUEST ##############################

  static Future<void> addGuest(String id, String name, String institute) async {
    try {
      final gsheets = GSheets(_credentials);
      final ss = await gsheets.spreadsheet(_spreadsheetId);
      final sheet = ss.worksheetByTitle('guest_list');
      if (sheet != null) {
        final DateFormat formatter = DateFormat('dd/MM/yyyy - HH:mm');
        final String formattedDate = formatter.format(DateTime.now());
        final newRow = [id, name, institute, formattedDate];
        await sheet.values.appendRow(newRow);
      } else {
        throw ('Worksheet not found');
      }
    } catch (e) {
      print('Error adding guest: $e');
      throw ('Error adding guest: $e');
    }
  }

  //############################################## Update Activity ######################################

  // Check if guest exists for update
  static Future<Map<String, dynamic>> checkIfGuestExists(String id) async {
    try {
      final gsheets = GSheets(_credentials);
      final ss = await gsheets.spreadsheet(_spreadsheetId);
      final sheet = ss.worksheetByTitle('guest_list');
      if (sheet != null) {
        final idColumn = await sheet.values.columnByKey('ID') ?? [];
        if (idColumn.contains(id)) {
          final index = idColumn.indexOf(id);
          final names = await sheet.values.columnByKey('Name') ?? [];
          if (index >= 0 && index < names.length && names[index].isNotEmpty) {
            final name = names[index];
            return {'if_exists': true, 'name': name, 'index': index};
          }
        }
      }
    } catch (e) {
      throw ('Error checking if guest exists: $e');
    }
    return {'if_exists': false, 'name': 'none', 'index': 0};
  }

  // Update QR Data
  static Future<dynamic> updateQRData(
    String qrData,
    String selectedDay,
    String selectedActivity,
    Map<String, dynamic> result,
  ) async {
    try {
      final gsheets = GSheets(_credentials);
      final ss = await gsheets.spreadsheet(_spreadsheetId);
      final sheet = ss.worksheetByTitle(selectedDay);
      if (sheet != null) {
        final int index = result['index'];
        if (result['if_exists']) {
          final List<String>? nameColumn =
              await sheet.values.columnByKey('Name');
          final List<String>? headers = await sheet.values.row(1);
          final int nameIndex = headers?.indexOf('Name') ?? -1;
          final int checkInIndex = headers?.indexOf('Registration Date') ?? -1;
          final int dinnerIndex = headers?.indexOf('Dinner') ?? -1;

          final List<String>? checkInColumn =
              await sheet.values.columnByKey('Registration Date');
          final List<String>? dinnerColumn =
              await sheet.values.columnByKey('Dinner');

          try {
            if (nameColumn == null ||
                nameColumn.isEmpty ||
                index >= nameColumn.length ||
                nameColumn[index] == '') {
              final String name = result['name'];
              await sheet.values
                  .insertValue(name, column: nameIndex + 1, row: index + 2);
            }
            final DateFormat formatter = DateFormat('dd/MM/yyyy - HH:mm');
            final String formattedDate = formatter.format(DateTime.now());

            if (selectedActivity == 'check_in') {
              if (checkInColumn != null &&
                  checkInColumn.length > index &&
                  checkInColumn[index].isNotEmpty) {
                return 'User already checked-in';
              } else {
                await sheet.values.insertValue(formattedDate,
                    column: checkInIndex + 1, row: index + 2);
              }
            } else if (selectedActivity == 'dinner') {
              if (dinnerColumn != null &&
                  dinnerColumn.length > index &&
                  dinnerColumn[index].isNotEmpty) {
                return 'User already completed dinner';
              } else {
                await sheet.values.insertValue(formattedDate,
                    column: dinnerIndex + 1, row: index + 2);
              }
            }
            return 'Updated';
          } catch (e) {
            return e.toString();
          }
        } else {
          final guestListSheet = ss.worksheetByTitle('guest_list');
          final guestIdColumn = await guestListSheet?.values.columnByKey('ID');
          final guestNameColumn =
              await guestListSheet?.values.columnByKey('Name');
          if (guestIdColumn != null && guestNameColumn != null) {
            final guestIndex = guestIdColumn.indexOf(qrData);
            if (guestIndex != -1) {
              final guestName = guestNameColumn[guestIndex];
              final DateFormat formatter = DateFormat('dd/MM/yyyy - HH:mm');
              final String formattedDate = formatter.format(DateTime.now());
              final newRow = {
                'ID': qrData,
                'Name': guestName,
                'Registration Date':
                    selectedActivity == 'check_in' ? formattedDate : '',
                'Dinner': selectedActivity == 'dinner' ? formattedDate : '',
              };
              final List<dynamic> newRowList = newRow.values.toList();
              await sheet.values.appendRow(newRowList);
              return 'Updated';
            }
          }
        }
      } else {
        throw ('Worksheet not found');
      }
    } catch (e) {
      print('Error updating QR data: $e');
      return 'Error: $e';
    }
  }


}
