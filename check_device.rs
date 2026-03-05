use hidapi::HidApi;

fn main() {
    let api = HidApi::new().expect("Failed to create HID API");
    
    println!("All Lenovo keyboard devices found:");
    println!("{:-<80}", "");
    
    let mut found = false;
    for device in api.device_list() {
        if device.vendor_id() == 0x048d {
            found = true;
            println!("Vendor ID:  0x{:04x}", device.vendor_id());
            println!("Product ID: 0x{:04x}", device.product_id());
            println!("Usage Page: 0x{:04x}", device.usage_page());
            println!("Usage:      0x{:04x}", device.usage());
            println!("Path:       {:?}", device.path());
            println!("{:-<80}", "");
        }
    }
    
    if !found {
        println!("No Lenovo devices (0x048d) found!");
        println!("Make sure the keyboard is connected and drivers are installed.");
    }
}
