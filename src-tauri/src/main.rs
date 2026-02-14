// Этот файл является частью time-to-table
// SPDX-License-Identifier: GPL-3.0-or-later

// Предотвращает появление дополнительного окна консоли в Windows при релизе, НЕ УДАЛЯТЬ!!
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

fn main() {
    time_sap_lib::run()
}
