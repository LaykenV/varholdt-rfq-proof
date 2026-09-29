# Synthetic RFQ proof

This archive contains original demonstration code and invented RFQ inputs.
Published as a synthetic demonstration for review.

Start with docs/rfq-quote.md for the exact supported PDF layout. Follow
docs/rfq-n8n.md to reproduce the complete local n8n path. The loopback service
is a test adapter. It is not a hosted customer upload service.

The baseline workbook has two priced rows totaling AED 25.02. The unseen
synthetic PDF has eight rows, two priced and six unpriced for review. Its
priced subtotal is AED 23.00. The rejected PDF returns no workbook.

All 49 repository tests passed. This archive contains the 17 RFQ tests,
including nine independent acceptance tests. The other 32 repository tests
cover earlier project work and are outside this archive.

Evidence folders include the actual workflow definitions, verification records,
workbooks and results. The verification records identify the exact source
hashes used. The n8n runtime, its database, credentials and diagnostic logs are
excluded. Install the pinned dependencies separately.

The parser handles only the documented synthetic text layout. Customer PDF
compatibility, Arabic-language review and desktop Excel visual inspection are
unverified. Amounts are static text. Human approval is required before a quote
is issued. No customer outcome, payment or production deployment is claimed.

The original source is supplied here for inspection. Client modification and
deployment rights remain proposed terms requiring owner approval and a written
agreement. Third-party dependencies retain their own licenses, linked in the
docs. This archive does not grant rights to the n8n runtime.
