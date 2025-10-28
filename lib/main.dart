import 'dart:convert';
import 'dart:io';
import 'dart:math' as math;

import 'package:desktop_drop/desktop_drop.dart';
import 'package:excel/excel.dart' as excel;
import 'package:flutter/material.dart';
import 'package:http/http.dart' as http;
import 'package:path/path.dart' as p;
import 'package:url_launcher/url_launcher.dart';
import 'package:window_size/window_size.dart' as window_size;
import 'package:xml/xml.dart';

const String kCurrentVersion = '1.0.0';
const String kGitHubOwner = 'jasmann';
const String kGitHubRepo = 'ftfix_flutter';

Future<void> main() async {
  WidgetsFlutterBinding.ensureInitialized();
  if (Platform.isWindows || Platform.isLinux || Platform.isMacOS) {
    window_size.setWindowTitle('FactoryTalk Alarm Bit Fixer');
    const minSize = Size(760, 520);
    const preferredSize = Size(900, 680);
    const maxSize = Size(4096, 4096);
    window_size.setWindowMinSize(minSize);
    window_size.setWindowMaxSize(maxSize);
    final screen = await window_size.getCurrentScreen();
    if (screen != null) {
      final frame = Rect.fromCenter(
        center: screen.visibleFrame.center,
        width: preferredSize.width,
        height: preferredSize.height,
      );
      window_size.setWindowFrame(frame);
    }
  }
  runApp(const FtFixApp());
}

class FtFixApp extends StatelessWidget {
  const FtFixApp({super.key});

  @override
  Widget build(BuildContext context) {
    final baseColorScheme = ColorScheme.fromSeed(
      seedColor: const Color(0xFF5865F2),
      brightness: Brightness.dark,
    );

    return MaterialApp(
      title: 'FT Alarm Bit Fix',
      theme: ThemeData(
        brightness: Brightness.dark,
        colorScheme: baseColorScheme.copyWith(
          surface: const Color(0xFF2B2D31),
          primary: const Color(0xFF5865F2),
          secondary: const Color(0xFF5865F2),
          outline: const Color(0xFF4E5058),
        ),
        scaffoldBackgroundColor: const Color(0xFF313338),
        appBarTheme: const AppBarTheme(
          backgroundColor: Color(0xFF2B2D31),
          foregroundColor: Colors.white,
          elevation: 0,
        ),
        elevatedButtonTheme: ElevatedButtonThemeData(
          style: ElevatedButton.styleFrom(
            backgroundColor: const Color(0xFF5865F2),
            foregroundColor: Colors.white,
            padding: const EdgeInsets.symmetric(horizontal: 20, vertical: 14),
            shape: RoundedRectangleBorder(
              borderRadius: BorderRadius.circular(10),
            ),
          ),
        ),
        cardColor: const Color(0xFF2B2D31),
        useMaterial3: true,
      ),
      home: const AlarmFixHome(),
    );
  }
}

class AlarmFixHome extends StatefulWidget {
  const AlarmFixHome({super.key});

  @override
  State<AlarmFixHome> createState() => _AlarmFixHomeState();
}

enum ExportFormat { csv, xlsx }

class _AlarmFixHomeState extends State<AlarmFixHome> {
  bool _isDragging = false;
  String? _sourcePath;
  String? _exportPath;
  String? _errorMessage;
  List<AlarmRecord> _records = const [];
  bool _isParsing = false;
  final ScrollController _verticalController = ScrollController();
  final ScrollController _horizontalController = ScrollController();
  bool _isCheckingUpdate = false;
  ReleaseInfo? _updateInfo;
  String? _updateError;

  @override
  void initState() {
    super.initState();
    WidgetsBinding.instance.addPostFrameCallback((_) => _checkForUpdates());
  }

  @override
  void dispose() {
    _verticalController.dispose();
    _horizontalController.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('FactoryTalk Alarm Bit Fixer'),
        actions: [
          IconButton(
            tooltip: 'About',
            icon: const Icon(Icons.info_outline),
            onPressed: _showAboutDialog,
          ),
          const SizedBox(width: 8),
        ],
      ),
      body: Padding(
        padding: const EdgeInsets.all(24),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Text(
              'Drop a FactoryTalk View SE/ME alarm export (*.xml) to shift the alarm bits to Studio 5000 style indexing.',
              style: Theme.of(context).textTheme.titleMedium,
            ),
            const SizedBox(height: 16),
            _buildDropZone(),
            const SizedBox(height: 12),
            _buildUpdateNotice(),
            const SizedBox(height: 16),
            if (_errorMessage != null)
              Text(
                _errorMessage!,
                style: TextStyle(color: Theme.of(context).colorScheme.error),
              ),
            if (_sourcePath != null)
              Padding(
                padding: const EdgeInsets.only(top: 4),
                child: Text(
                  'Loaded: $_sourcePath',
                  style: Theme.of(context).textTheme.bodySmall,
                ),
              ),
            if (_exportPath != null)
              Padding(
                padding: const EdgeInsets.only(top: 4),
                child: Text(
                  'Exported: $_exportPath',
                  style: Theme.of(context).textTheme.bodySmall,
                ),
              ),
            const SizedBox(height: 16),
            Flexible(
              child: _buildResults(),
            ),
            const SizedBox(height: 24),
            Wrap(
              spacing: 12,
              runSpacing: 12,
              alignment: WrapAlignment.end,
              children: [
                ElevatedButton.icon(
                  onPressed: _records.isEmpty || _isParsing
                      ? null
                      : () => _exportData(ExportFormat.csv),
                  icon: const Icon(Icons.save_alt),
                  label: const Text('Export CSV'),
                ),
                ElevatedButton.icon(
                  onPressed: _records.isEmpty || _isParsing
                      ? null
                      : () => _exportData(ExportFormat.xlsx),
                  icon: const Icon(Icons.grid_on),
                  label: const Text('Export XLSX'),
                ),
              ],
            ),
          ],
        ),
      ),
    );
  }

  Widget _buildDropZone() {
    final colorScheme = Theme.of(context).colorScheme;
    final borderColor =
        _isDragging ? colorScheme.primary : colorScheme.outline;

    return DropTarget(
      onDragEntered: (_) {
        setState(() {
          _isDragging = true;
        });
      },
      onDragExited: (_) {
        setState(() {
          _isDragging = false;
        });
      },
      onDragDone: (details) async {
        setState(() {
          _isDragging = false;
        });
        await _handleDrop(details);
      },
      child: AnimatedContainer(
        duration: const Duration(milliseconds: 120),
        height: 180,
        width: double.infinity,
        decoration: BoxDecoration(
          color: _isDragging
              ? colorScheme.primary.withValues(alpha: 0.05)
              : colorScheme.surfaceContainerHighest.withValues(alpha: 0.4),
          borderRadius: BorderRadius.circular(16),
          border: Border.all(color: borderColor, width: 2),
        ),
        child: Center(
          child: _isParsing
              ? const CircularProgressIndicator()
              : Column(
                  mainAxisAlignment: MainAxisAlignment.center,
                  children: [
                    Icon(
                      Icons.file_upload,
                      size: 40,
                      color: borderColor,
                    ),
                    const SizedBox(height: 12),
                    Text(
                      'Drag & drop an alarm export XML file here',
                      style: Theme.of(context).textTheme.titleMedium,
                    ),
                    const SizedBox(height: 4),
                    Text(
                      'We will shift only the trailing bit number (e.g. Tag[1].1 -> Tag[1].0).',
                      style: Theme.of(context).textTheme.bodySmall,
                      textAlign: TextAlign.center,
                    ),
                  ],
                ),
        ),
      ),
    );
  }

  Widget _buildUpdateNotice() {
    if (_isCheckingUpdate) {
      return const SizedBox.shrink();
    }

    final colorScheme = Theme.of(context).colorScheme;

    if (_updateInfo?.isNewer == true) {
      final info = _updateInfo!;
      return Card(
        color: colorScheme.primary.withValues(alpha: 0.15),
        child: ListTile(
          leading: Icon(
            Icons.system_update_alt,
            color: colorScheme.primary,
          ),
          title: Text('Update available: v${info.version}'),
          subtitle: const Text('A newer release is ready on GitHub.'),
          trailing: TextButton(
            onPressed: _launchUpdateUrl,
            child: const Text('Get Update'),
          ),
        ),
      );
    }

    if (_updateError != null) {
      return Card(
        color: colorScheme.error.withValues(alpha: 0.15),
        child: ListTile(
          leading: Icon(
            Icons.error_outline,
            color: colorScheme.error,
          ),
          title: const Text('Unable to check for updates'),
          subtitle: Text(_updateError!),
          trailing: TextButton(
            onPressed: _isCheckingUpdate ? null : _checkForUpdates,
            child: const Text('Retry'),
          ),
        ),
      );
    }

    return const SizedBox.shrink();
  }

  Widget _buildResults() {
    if (_records.isEmpty) {
      return const Center(
        child: Text(
          'Results will appear here after you drop an XML file.',
          textAlign: TextAlign.center,
        ),
      );
    }

    return LayoutBuilder(
      builder: (context, constraints) {
        return Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Text(
              'Records: ${_records.length}',
              style: Theme.of(context).textTheme.titleSmall,
            ),
            const SizedBox(height: 8),
            Flexible(
              child: Scrollbar(
                controller: _verticalController,
                thumbVisibility: true,
                child: SingleChildScrollView(
                  controller: _verticalController,
                  scrollDirection: Axis.vertical,
                  child: Scrollbar(
                    controller: _horizontalController,
                    thumbVisibility: true,
                    notificationPredicate: (notification) =>
                        notification.metrics.axis == Axis.horizontal,
                    child: SingleChildScrollView(
                      controller: _horizontalController,
                      scrollDirection: Axis.horizontal,
                      child: ConstrainedBox(
                        constraints: BoxConstraints(
                          minWidth: constraints.maxWidth,
                        ),
                        child: DataTable(
                          columns: const [
                            DataColumn(label: Text('Studio 5000 Address')),
                            DataColumn(label: Text('Alarm Text')),
                          ],
                          rows: _records
                              .map(
                                (record) => DataRow(
                                  cells: [
                                    DataCell(Text(record.correctedAddress)),
                                    DataCell(
                                      Text(
                                        record.messageText,
                                        softWrap: false,
                                      ),
                                    ),
                                  ],
                                ),
                              )
                              .toList(),
                        ),
                      ),
                    ),
                  ),
                ),
              ),
            ),
          ],
        );
      },
    );
  }

  Future<void> _checkForUpdates() async {
    if (_isCheckingUpdate || Platform.environment.containsKey('FLUTTER_TEST')) {
      return;
    }

    setState(() {
      _isCheckingUpdate = true;
      _updateError = null;
    });

    ReleaseInfo? info;
    String? error;

    try {
      info = await _fetchLatestRelease();
    } catch (_) {
      error = 'Check your connection and try again.';
    }

    if (!mounted) {
      return;
    }

    setState(() {
      _isCheckingUpdate = false;
      _updateInfo = info;
      _updateError = error;
    });
  }

  Future<ReleaseInfo?> _fetchLatestRelease() async {
    final uri = Uri.https(
      'api.github.com',
      '/repos/$kGitHubOwner/$kGitHubRepo/releases/latest',
    );

    final response = await http.get(
      uri,
      headers: const {
        'Accept': 'application/vnd.github+json',
        'User-Agent': 'ftfix_flutter/$kCurrentVersion',
      },
    );

    if (response.statusCode != 200) {
      throw HttpException(
        'GitHub returned ${response.statusCode}',
        uri: uri,
      );
    }

    final payload = jsonDecode(response.body);
    if (payload is! Map<String, dynamic>) {
      return null;
    }

    final rawTag =
        (payload['tag_name'] ?? payload['name'] ?? '').toString().trim();
    final version = _normalizeReleaseVersion(rawTag);
    if (version == null) {
      return null;
    }

    final htmlUrlString = payload['html_url'] as String?;
    final htmlUrl = htmlUrlString != null && htmlUrlString.isNotEmpty
        ? Uri.parse(htmlUrlString)
        : Uri.parse('https://github.com/$kGitHubOwner/$kGitHubRepo/releases');

    final isNewer = _isVersionNewer(version, kCurrentVersion);
    return ReleaseInfo(
      version: version,
      htmlUrl: htmlUrl,
      isNewer: isNewer,
    );
  }

  String? _normalizeReleaseVersion(String raw) {
    if (raw.isEmpty) {
      return null;
    }
    final trimmed = raw.trim();
    if (trimmed.isEmpty) {
      return null;
    }
    final withoutPrefix =
        trimmed.startsWith(RegExp(r'[vV]')) ? trimmed.substring(1) : trimmed;
    return withoutPrefix.isEmpty ? null : withoutPrefix;
  }

  bool _isVersionNewer(String latest, String current) {
    final latestParts = _parseVersionParts(latest);
    final currentParts = _parseVersionParts(current);
    if (latestParts == null || currentParts == null) {
      return latest.compareTo(current) > 0;
    }

    final length = math.max(latestParts.length, currentParts.length);
    for (var i = 0; i < length; i++) {
      final latestValue = i < latestParts.length ? latestParts[i] : 0;
      final currentValue = i < currentParts.length ? currentParts[i] : 0;
      if (latestValue > currentValue) {
        return true;
      }
      if (latestValue < currentValue) {
        return false;
      }
    }

    return false;
  }

  List<int>? _parseVersionParts(String version) {
    var sanitized = version.trim();
    if (sanitized.isEmpty) {
      return null;
    }
    if (sanitized.startsWith(RegExp(r'[vV]'))) {
      sanitized = sanitized.substring(1);
    }
    sanitized = sanitized.split(RegExp(r'[-+]')).first;
    if (sanitized.isEmpty) {
      return null;
    }
    final segments = sanitized.split('.');
    if (segments.isEmpty) {
      return null;
    }
    final values = <int>[];
    for (final segment in segments) {
      if (segment.isEmpty) {
        return null;
      }
      final value = int.tryParse(segment);
      if (value == null) {
        return null;
      }
      values.add(value);
    }
    return values;
  }

  Future<void> _launchUpdateUrl() async {
    final info = _updateInfo;
    if (info == null) {
      return;
    }

    final launched = await launchUrl(
      info.htmlUrl,
      mode: LaunchMode.externalApplication,
    );

    if (!launched && mounted) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(
          content: Text('Could not open the release page.'),
        ),
      );
    }
  }

  Future<void> _handleDrop(DropDoneDetails details) async {
    if (details.files.isEmpty) {
      return;
    }

    DropItem? selected;
    for (final item in details.files) {
      if (item is DropItemDirectory) {
        continue;
      }
      selected ??= item;
      final mime = item.mimeType?.toLowerCase() ?? '';
      final name = item.name.toLowerCase();
      if (mime.contains('xml') || name.endsWith('.xml')) {
        selected = item;
        break;
      }
    }

    if (selected == null) {
      setState(() {
        _errorMessage = 'No file detected in the dropped items.';
      });
      return;
    }

    final path = selected.path;
    if (path.isEmpty) {
      setState(() {
        _errorMessage = 'Could not read the dropped item path.';
      });
      return;
    }

    setState(() {
      _isParsing = true;
      _errorMessage = null;
      _exportPath = null;
      _records = const [];
      _sourcePath = path;
    });

    try {
      final file = File(path);
      if (!await file.exists()) {
        throw FileSystemException('File not found', path);
      }
      final content = await file.readAsString();
      final document = XmlDocument.parse(content);
      final records = _parseDocument(document);

      setState(() {
        _records = records;
      });
    } on XmlParserException catch (error) {
      setState(() {
        _errorMessage = 'Invalid XML file: ${error.message}';
      });
    } on FormatException catch (error) {
      setState(() {
        _errorMessage = error.message;
      });
    } on FileSystemException catch (error) {
      setState(() {
        _errorMessage =
            'Unable to read the file (${error.path ?? 'unknown path'}).';
      });
    } catch (error) {
      setState(() {
        _errorMessage = 'Unexpected error: $error';
      });
    } finally {
      setState(() {
        _isParsing = false;
      });
    }
  }

  List<AlarmRecord> _parseDocument(XmlDocument document) {
    final triggerMap = <String, TriggerInfo>{};

    for (final trigger in document.findAllElements('trigger')) {
      final id = trigger.getAttribute('id');
      final exp = trigger.getAttribute('exp');
      if (id == null || exp == null) {
        continue;
      }
      triggerMap[id] = TriggerInfo(
        id: id,
        baseTag: _normalizeExpression(exp),
        label: trigger.getAttribute('label') ?? '',
      );
    }

    final records = <AlarmRecord>[];

    for (final message in document.findAllElements('message')) {
      final triggerRef = message.getAttribute('trigger');
      final valueRaw = message.getAttribute('trigger-value');
      if (triggerRef == null || valueRaw == null) {
        continue;
      }

      final triggerId = triggerRef.startsWith('#')
          ? triggerRef.substring(1)
          : triggerRef;
      final triggerInfo = triggerMap[triggerId];
      if (triggerInfo == null) {
        continue;
      }

      final value = int.tryParse(valueRaw);
      if (value == null) {
        continue;
      }

      if (value <= 0) {
        // Skip invalid bit values; FactoryTalk exports should start at 1.
        continue;
      }

      final corrected = value - 1;
      final baseTag = triggerInfo.baseTag;
      final originalAddress = '$baseTag.$value';
      final correctedAddress = '$baseTag.$corrected';
      final messageText = message.getAttribute('text') ?? '';

      records.add(
        AlarmRecord(
          triggerId: triggerId,
          triggerLabel: triggerInfo.label,
          baseTag: baseTag,
          originalBit: value,
          correctedBit: corrected,
          originalAddress: originalAddress,
          correctedAddress: correctedAddress,
          messageText: messageText,
        ),
      );
    }

    return records;
  }

  Future<void> _exportData(ExportFormat format) async {
    final source = _sourcePath;
    if (source == null || _records.isEmpty) {
      return;
    }

    final directory = p.dirname(source);
    final baseName = p.basenameWithoutExtension(source);
    final exportName = format == ExportFormat.csv
        ? '${baseName}_corrected.csv'
        : '${baseName}_corrected.xlsx';
    final file = File(p.join(directory, exportName));

    if (format == ExportFormat.csv) {
      await file.writeAsString(_buildCsvContent());
    } else {
      await file.writeAsBytes(await _buildXlsxBytes(), flush: true);
    }

    setState(() {
      _exportPath = file.path;
    });
  }

  String _buildCsvContent() {
    final buffer = StringBuffer()
      ..writeln(
        _csvLine([
          'Studio 5000 Address',
          'Alarm Text',
        ]),
      );

    for (final record in _records) {
      buffer.writeln(
        _csvLine([
          record.correctedAddress,
          record.messageText,
        ]),
      );
    }

    return buffer.toString();
  }

  Future<List<int>> _buildXlsxBytes() async {
    final workbook = excel.Excel.createExcel();
    final sheet = workbook['Sheet1'];
    sheet.appendRow([
      excel.TextCellValue('Studio 5000 Address'),
      excel.TextCellValue('Alarm Text'),
    ]);

    for (final record in _records) {
      sheet.appendRow([
        excel.TextCellValue(record.correctedAddress),
        excel.TextCellValue(record.messageText),
      ]);
    }

    final bytes = workbook.encode();
    if (bytes == null) {
      throw const FormatException('Failed to encode XLSX file.');
    }
    return bytes;
  }

  String _csvLine(List<String> values) {
    return values.map(_escapeCsvField).join(',');
  }

  String _escapeCsvField(String value) {
    final needsQuoting =
        value.contains(',') || value.contains('"') || value.contains('\n');
    final escaped = value.replaceAll('"', '""');
    return needsQuoting ? '"$escaped"' : escaped;
  }

  String _normalizeExpression(String expression) {
    var result = expression.trim();
    if (result.startsWith('{') && result.endsWith('}')) {
      result = result.substring(1, result.length - 1);
    }
    result = result.replaceAll('[PLC]', '');
    return result;
  }

  void _showAboutDialog() {
    showAboutDialog(
      context: context,
      applicationName: 'FT Alarm Bit Fix',
      applicationVersion: kCurrentVersion,
      applicationIcon: SizedBox(
        width: 48,
        height: 48,
        child: DecoratedBox(
          decoration: const BoxDecoration(
            color: Color(0xFF5865F2),
            shape: BoxShape.circle,
          ),
          child: const Center(
            child: Text(
              'FT',
              style: TextStyle(
                color: Colors.white,
                fontWeight: FontWeight.bold,
              ),
            ),
          ),
        ),
      ),
      children: const [
        SizedBox(height: 12),
        Text(
          'FactoryTalk Alarm Bit Fix shifts SE/ME alarm exports to the 0-based '
          'Studio 5000 addressing scheme, keeping descriptions intact and ready '
          'for CSV/XLSX import.',
        ),
        SizedBox(height: 12),
        Text(
          'Drop an FT View XML file into the window, review the corrected tags, '
          'and export the format your team needs.',
        ),
      ],
    );
  }
}

class TriggerInfo {
  TriggerInfo({
    required this.id,
    required this.baseTag,
    required this.label,
  });

  final String id;
  final String baseTag;
  final String label;
}

class AlarmRecord {
  const AlarmRecord({
    required this.triggerId,
    required this.triggerLabel,
    required this.baseTag,
    required this.originalBit,
    required this.correctedBit,
    required this.originalAddress,
    required this.correctedAddress,
    required this.messageText,
  });

  final String triggerId;
  final String triggerLabel;
  final String baseTag;
  final int originalBit;
  final int correctedBit;
  final String originalAddress;
  final String correctedAddress;
  final String messageText;

  String get triggerDisplay =>
      triggerLabel.isNotEmpty ? '$triggerId ($triggerLabel)' : triggerId;
}

class ReleaseInfo {
  const ReleaseInfo({
    required this.version,
    required this.htmlUrl,
    required this.isNewer,
  });

  final String version;
  final Uri htmlUrl;
  final bool isNewer;
}
