#!/bin/bash

# Fetch MX, SPF, DKIM, and DMARC DNS records for a domain on macOS.
# Uses dig, which is available on macOS.
#
# Examples:
#   ./get-mail-dns-records-mac.sh example.com
#   ./get-mail-dns-records-mac.sh -d example.com
#   ./get-mail-dns-records-mac.sh -d example.com -s 8.8.8.8
#   ./get-mail-dns-records-mac.sh -d example.com -k selector1,selector2,google

DEFAULT_SELECTORS=(
  "selector1"
  "selector2"
  "google"
  "default"
  "dkim"
  "mail"
  "smtp"
  "s1"
  "s2"
  "k1"
  "k2"
)

DOMAIN=""
DNS_SERVER=""
DKIM_SELECTORS=("${DEFAULT_SELECTORS[@]}")

show_usage() {
  cat <<EOF
Usage:
  $0 <domain>
  $0 -d <domain> [options]

Options:
  -d, --domain DOMAIN              Domain to check
  -s, --server DNS_SERVER          DNS resolver to use, for example 8.8.8.8 or 1.1.1.1
  -k, --dkim-selectors SELECTORS   Comma-separated DKIM selectors to check
  -h, --help                       Show this help message

Examples:
  $0 example.com
  $0 -d example.com -s 8.8.8.8
  $0 -d example.com -k selector1,selector2,google
EOF
}

die() {
  echo "Error: $*" >&2
  exit 1
}

trim() {
  printf "%s" "$1" | sed 's/^[[:space:]]*//; s/[[:space:]]*$//'
}

normalize_domain() {
  local input="$1"
  local clean

  clean="$(trim "$input" | tr '[:upper:]' '[:lower:]')"
  clean="${clean#http://}"
  clean="${clean#https://}"
  clean="${clean%%/*}"
  clean="${clean%%:*}"
  clean="${clean%.}"

  printf "%s" "$clean"
}

split_selectors() {
  local selector_string="$1"
  local old_ifs="$IFS"

  IFS=","
  read -r -a DKIM_SELECTORS <<< "$selector_string"
  IFS="$old_ifs"
}

while [ "$#" -gt 0 ]; do
  case "$1" in
    -d|--domain)
      shift
      [ "$#" -gt 0 ] || die "Missing value for --domain"
      DOMAIN="$1"
      ;;
    -s|--server|--dns-server|--resolver)
      shift
      [ "$#" -gt 0 ] || die "Missing value for --server"
      DNS_SERVER="$1"
      ;;
    -k|--dkim-selectors|--selectors)
      shift
      [ "$#" -gt 0 ] || die "Missing value for --dkim-selectors"
      split_selectors "$1"
      ;;
    -h|--help)
      show_usage
      exit 0
      ;;
    -*)
      die "Unknown option: $1"
      ;;
    *)
      if [ -z "$DOMAIN" ]; then
        DOMAIN="$1"
      else
        die "Unexpected argument: $1"
      fi
      ;;
  esac

  shift
done

[ -n "$DOMAIN" ] || {
  show_usage
  exit 1
}

DOMAIN="$(normalize_domain "$DOMAIN")"

[ -n "$DOMAIN" ] || die "Domain could not be parsed."

command -v dig >/dev/null 2>&1 || die "dig was not found. This script requires dig."

dig_query() {
  local name="$1"
  local type="$2"

  if [ -n "$DNS_SERVER" ]; then
    dig @"$DNS_SERVER" "$name" "$type" +short +time=3 +tries=2
  else
    dig "$name" "$type" +short +time=3 +tries=2
  fi
}

txt_line_to_value() {
  local line="$1"

  # dig prints split TXT records like:
  #   "part one" "part two"
  # DNS clients concatenate these TXT chunks. This removes the boundaries.
  printf "%s\n" "$line" |
    sed \
      -e 's/"[[:space:]]*"//g' \
      -e 's/^"//' \
      -e 's/"$//' \
      -e 's/\\"/"/g'
}

get_txt_records() {
  local name="$1"
  local line

  dig_query "$name" TXT | while IFS= read -r line; do
    case "$line" in
      \"*)
        txt_line_to_value "$line"
        ;;
    esac
  done
}

get_cname_record() {
  local name="$1"

  dig_query "$name" CNAME |
    head -n 1 |
    sed 's/\.$//'
}

count_non_empty_lines() {
  local input="$1"

  if [ -z "$input" ]; then
    echo 0
  else
    printf "%s\n" "$input" |
      sed '/^[[:space:]]*$/d' |
      wc -l |
      tr -d ' '
  fi
}

get_dkim_query_name() {
  local selector="$1"
  local base_domain="$2"

  selector="$(trim "$selector")"
  selector="${selector%.}"

  if echo "$selector" | grep -Eq '\._domainkey\.'; then
    printf "%s" "$selector"
  elif echo "$selector" | grep -Eq '\._domainkey$'; then
    printf "%s.%s" "$selector" "$base_domain"
  else
    printf "%s._domainkey.%s" "$selector" "$base_domain"
  fi
}

print_section() {
  echo ""
  echo "$1"
  echo "$2"
}

DNS_DISPLAY="System default"
[ -n "$DNS_SERVER" ] && DNS_DISPLAY="$DNS_SERVER"

echo ""
echo "Domain:    $DOMAIN"
echo "DNS:       $DNS_DISPLAY"
echo "Queried:   $(date)"
echo ""

print_section "MX Records" "----------"

MX_RECORDS="$(dig_query "$DOMAIN" MX | sort -n | sed 's/\.$//')"

if [ -n "$MX_RECORDS" ]; then
  printf "%s\n" "$MX_RECORDS" | while read -r preference exchange rest; do
    if [ -n "$preference" ] && [ -n "$exchange" ]; then
      printf "Preference: %-6s Exchange: %s\n" "$preference" "$exchange"
    fi
  done
else
  echo "No MX records found."
fi

print_section "SPF Records" "-----------"

DOMAIN_TXT_RECORDS="$(get_txt_records "$DOMAIN")"

SPF_RECORDS="$(
  printf "%s\n" "$DOMAIN_TXT_RECORDS" |
    grep -Ei '^[[:space:]]*v=spf1([[:space:]]|$)'
)"

SPF_COUNT="$(count_non_empty_lines "$SPF_RECORDS")"

if [ "$SPF_COUNT" -gt 0 ]; then
  printf "%s\n" "$SPF_RECORDS" | while IFS= read -r record; do
    [ -n "$record" ] && echo "$record"
  done
else
  echo "No SPF record found at $DOMAIN."
fi

print_section "DMARC Records" "-------------"

DMARC_NAME="_dmarc.$DOMAIN"

DMARC_RECORDS="$(
  get_txt_records "$DMARC_NAME" |
    grep -Ei '^[[:space:]]*v=DMARC1([[:space:]]*;|$)'
)"

DMARC_COUNT="$(count_non_empty_lines "$DMARC_RECORDS")"

if [ "$DMARC_COUNT" -gt 0 ]; then
  printf "%s\n" "$DMARC_RECORDS" | while IFS= read -r record; do
    [ -n "$record" ] && echo "$record"
  done
else
  echo "No DMARC record found at $DMARC_NAME."
fi

print_section "DKIM Records" "------------"

DKIM_CHECKED=0

for selector in "${DKIM_SELECTORS[@]}"; do
  selector="$(trim "$selector")"
  [ -n "$selector" ] || continue

  DKIM_CHECKED=$((DKIM_CHECKED + 1))

  QUERY_NAME="$(get_dkim_query_name "$selector" "$DOMAIN")"
  DKIM_TXT_RECORDS="$(get_txt_records "$QUERY_NAME")"
  CNAME_TARGET="$(get_cname_record "$QUERY_NAME")"

  if [ -z "$DKIM_TXT_RECORDS" ] && [ -n "$CNAME_TARGET" ]; then
    DKIM_TXT_RECORDS="$(get_txt_records "$CNAME_TARGET")"
  fi

  DKIM_COUNT="$(count_non_empty_lines "$DKIM_TXT_RECORDS")"

  if [ "$DKIM_COUNT" -eq 0 ]; then
    STATUS="Not found"
  elif printf "%s\n" "$DKIM_TXT_RECORDS" | grep -Eiq '(^[[:space:]]*v=DKIM1([[:space:];]|$)|;[[:space:]]*p=|^[[:space:]]*k=rsa[[:space:]]*;)'; then
    STATUS="Found"
  else
    STATUS="TXT found; verify manually"
  fi

  echo "Selector:  $selector"
  echo "Query:     $QUERY_NAME"
  echo "Status:    $STATUS"

  if [ -n "$CNAME_TARGET" ]; then
    echo "CNAME:     $CNAME_TARGET"
  fi

  if [ "$DKIM_COUNT" -gt 0 ]; then
    echo "TXT:"
    printf "%s\n" "$DKIM_TXT_RECORDS" | while IFS= read -r record; do
      [ -n "$record" ] && echo "  $record"
    done
  fi

  echo ""
done

print_section "Warnings" "--------"

WARNINGS_FOUND=0

if [ "$SPF_COUNT" -gt 1 ]; then
  echo "- Multiple SPF records found. A domain should normally publish exactly one SPF TXT record."
  WARNINGS_FOUND=1
fi

if [ "$DMARC_COUNT" -gt 1 ]; then
  echo "- Multiple DMARC records found. A domain should normally publish exactly one DMARC TXT record at $DMARC_NAME."
  WARNINGS_FOUND=1
fi

if [ "$DKIM_CHECKED" -eq 0 ]; then
  echo "- No DKIM selectors were checked."
  WARNINGS_FOUND=1
fi

if [ "$WARNINGS_FOUND" -eq 0 ]; then
  echo "No warnings."
fi

echo ""