# IEABank October 2026 — User GUI migration checklist

Source: To-do Item Bank 20261008; student_v3_teacher_v2 workbook. This is preparation, NOT a production deploy.

## Data API contract (new workbook)
admin.instrument_pop, instrument_type, instrument_name, target_grade;
topic(topic_id,category_name); item(item_uid,item_name,topic_id,is_test_item);
item_admin(response_type,is_public_itembank,is_puf,is_version_published,wording_question_add);
response_scheme(is_checkbox); response_scheme_values(response_value,response_label,checked_value,not_checked_value);
index(index_uid,index_name,is_scaled); index_version(index_id,index_version_description,index_varname,is_public_itembank,is_puf,index_uid,index_var); index_item(index_id,index_item_pk,item_admin_pk).

## Public Explorer acceptance criteria
- Tabs Home, Items, Topics, Derived Variables, Instruments, Export Selection, Contact; add glossary/tutorial/news placeholders.
- Selection cart text -> Export Selection or Selected Information (final copy TBD).
- All visible 'scale' terminology -> 'derived variable'; 'item variation' -> 'item version'; 'category' -> 'topic'.
- Search must match item ID like IAA003V01 as well as item UID, wording, variables, topics.
- Item sidebar: instrument_name, instrument_pop, target_grade, topic; remove cycle.
- Display wording_question + wording_question_add.
- Public visibility: is_public_itembank; when is_puf=FALSE show RUF notice and repository URL.
- Multiple response checkbox scheme: append checked/not checked codings to labels.
- Derived Variables: is_scaled filter; composition over time; clickable version IDs and item names.
- Homepage layout and IEA colors #d91a15 #fdda6b #0071b9 #55585a.
- Export page redesign marked AC; leave for follow-up clarification.

## Blocking existing-code coupling
- app.R calls table 'scale_items' (plural), but current database is 'scale_item'; verify actual runtime and replace with new 'index_item'.
- app.R uses old scale_* and item_category columns; replace data assembly and filters together.
- Never deploy main until new DB + public RLS checks + test fixtures are validated.
