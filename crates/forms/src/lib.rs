//! Nelson's VB6 forms, read so the Rust ports can lay out their screens
//! from his designs: [`frm`] parses the control tree, [`frx`] reads the
//! values kept in the binary resource file next to it.

pub mod frm;
pub mod frx;

pub use frm::{Control, Font, Form, Value};
