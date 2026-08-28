# Doppio

VBA modules for connecting Excel to **Infor M3** via the MI API, ION API, IPS/SOAP web services, and IDM (Infor Document Management).

Built by [Doppio Group](mailto:eric@doppiogroup.com) · v2.01

---

## Modules

| Module | Description |
|--------|-------------|
| `DoppioCore` | Core types, enums, constants, and shared utilities |
| `DoppioConfig` | Environment configuration and tenant management |
| `DoppioAuth` | Authentication — tenant token, service account, token refresh |
| `DoppioHttp` | HTTP request execution and response handling |
| `DoppioCache` | In-memory and sheet-backed response caching |
| `DoppioProcess` | MI bulk call execution and result processing |
| `DoppioExportMI` | Export pipeline for MI program data |
| `DoppioBridge` | Backward-compatible bridge to legacy `apicall` interface |
| `DoppioUI` | UI helpers and sheet interactions |
| `DoppioApi` | High-level API entry points |
| `Doppio` | Main module — environment setup, layout, transactions, logging |
| `DoppioTest` | Test suite for auth, API calls, and migration validation |
| `JsonConverter` | JSON parsing ([VBA-JSON](https://github.com/VBA-tools/VBA-JSON)) |

## API Types Supported

- **MI** — M3 Management Interface programs
- **IDM** — Infor Document Management
- **IPS** — IPS/SOAP web services
- **Swagger / M3X** — OpenAPI and M3X format endpoints

## Debug Mode

All `Debug.Print` statements are gated by a compiler constant in `DoppioCore.bas`:

```vba
#Const DEBUG_MODE = True   ' set to False to silence all debug output
```

Output appears in the VBA **Immediate Window** when enabled.

## Usage

1. Import the `.bas` files into your Excel workbook via the VBA editor (Alt+F11 → File → Import)
2. Set your environment on the designated sheet (`Environment` named range)
3. Call `Doppio.Tenant_Token` to authenticate, then use `DoppioProcess` or `DoppioExportMI` to fetch data

## Requirements

- Microsoft Excel (Windows or Mac) with VBA enabled
- Infor M3 / CloudSuite environment with API access
- ION API credentials (client ID, client secret, tenant ID)
