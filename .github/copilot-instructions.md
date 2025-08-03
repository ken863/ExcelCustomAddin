---
applyTo: "**"
---
# Project general coding standards

## General coding standards
- Tất cả mã nguồn phải tuân thủ các tiêu chuẩn mã hóa hiện đại và tốt nhất.
- Sử dụng các công cụ kiểm tra mã nguồn như ESLint và Prettier để đảm bảo mã nguồn sạch sẽ và nhất quán.
- Sử dụng `git` để quản lý mã nguồn, commit thường xuyên và có ý nghĩa.
- Tất cả các commit phải có thông điệp rõ ràng, mô tả ngắn gọn về thay đổi đã thực hiện.
- Sử dụng `git branch` để quản lý các tính năng và sửa lỗi, không commit trực tiếp vào nhánh `main` hoặc `master`. 
- Tất cả các pull request phải được xem xét và phê duyệt bởi ít nhất một người khác trước khi được hợp nhất vào nhánh chính.
- Sử dụng `semantic versioning` cho các phiên bản của dự án, bao gồm các tag rõ ràng cho mỗi phiên bản.
- Tất cả các tệp mã nguồn phải được đặt trong thư mục `src` và không được để lẫn lộn với các tệp cấu hình hoặc tài liệu khác.
- Tất cả các tệp cấu hình phải được đặt trong thư mục `config` và không được để lẫn lộn với mã nguồn.
- Tất cả các tệp tài liệu phải được đặt trong thư mục `docs` và không được để lẫn lộn với mã nguồn hoặc cấu hình.
- Tất cả các tệp tài nguyên tĩnh (như hình ảnh, biểu tượng, v.v.) phải được đặt trong thư mục `assets` và không được để lẫn lộn với mã nguồn hoặc cấu hình.
- Tất cả các tệp liên quan đến Docker phải được đặt trong thư mục `Docker` và không được để lẫn lộn với mã nguồn hoặc cấu hình.
- Tất cả các tệp liên quan đến CI/CD phải được đặt trong thư mục `.github/workflows` và không được để lẫn lộn với mã nguồn hoặc cấu hình.
- Tất cả các tệp liên quan đến kiểm thử phải được đặt trong thư mục `tests` và không được để lẫn lộn với mã nguồn hoặc cấu hình.
- Tất cả các tệp liên quan đến tài liệu API phải được đặt trong thư mục `api-docs` và không được để lẫn lộn với mã nguồn hoặc cấu hình.
- Tất cả các tệp liên quan đến tài liệu hướng dẫn sử dụng phải được đặt trong thư mục `guides` và không được để lẫn lộn với mã nguồn hoặc cấu hình.
- Tất cả các tệp liên quan đến tài liệu kiến trúc phải được đặt trong thư mục `architecture` và không được để lẫn lộn với mã nguồn hoặc cấu hình.
- Tất cả các tệp liên quan đến tài liệu thiết kế phải được đặt trong thư mục `design` và không được để lẫn lộn với mã nguồn hoặc cấu hình.
- Tất cả các tệp liên quan đến tài liệu triển khai phải được đặt trong thư mục `deployment` và không được để lẫn lộn với mã nguồn hoặc cấu hình.
- Tất cả các tệp liên quan đến tài liệu bảo trì phải được đặt trong thư mục `maintenance` và không được để lẫn lộn với mã nguồn hoặc cấu hình.
- Tất cả các tệp liên quan đến tài liệu bảo mật phải được đặt trong thư mục `security` và không được để lẫn lộn với mã nguồn hoặc cấu hình.
- Tất cả các tệp liên quan đến tài liệu tuân thủ phải được đặt trong thư mục `compliance` và không được để lẫn lộn với mã nguồn hoặc cấu hình.

## Reponsibility
- Không cần tóm tắt các thay đổi khi phản hồi yêu cầu

## Response language
- Use Vietnamese for all code comments, documentation, and commit messages

## readme.md
- Tất cả các tệp README.md phải được viết bằng tiếng Việt, giải thích rõ ràng về cách sử dụng, cấu hình và triển khai dự án.
- Chỉ tạo 1 file .md duy nhất trong thư mục gốc của dự án, không tạo các tệp .md khác trong các thư mục con.

## React và JavaScript
- Sử dụng React cho giao diện người dùng, không sử dụng các framework khác như Angular hoặc Vue.js.
- Sử dụng JavaScript ES6+ cho mã nguồn, không sử dụng các phiên bản cũ hơn.
- Sử dụng `const` và `let` thay vì `var` để khai báo biến.
- Sử dụng arrow functions thay vì function expressions truyền thống.
- Sử dụng destructuring để truy cập các thuộc tính của đối tượng.
- Sử dụng template literals thay vì chuỗi nối.              
- Sử dụng `async/await` thay vì callback hoặc Promise chaining để xử lý bất đồng bộ.
- Sử dụng `import` và `export` thay vì `require` và `module.exports` để quản lý module.
- Sử dụng `class` thay vì prototype-based inheritance cho các component React.
- Sử dụng `useState`, `useEffect` và các hook khác để quản lý state và lifecycle trong component function.
- Sử dụng `PropTypes` hoặc TypeScript để xác định kiểu dữ liệu của props trong component.
- Sử dụng `styled-components` để tạo kiểu cho các component React.
- Sử dụng `axios` hoặc `fetch` để thực hiện các yêu cầu HTTP, không sử dụng jQuery.
- Sử dụng `ESLint` và `Prettier` để kiểm tra và định dạng mã nguồn.
- Sử dụng `Jest` và `React Testing Library` để viết và chạy các bài kiểm tra đơn vị cho component React.
- Sử dụng `Redux` hoặc `Context API` để quản lý state toàn cục, không sử dụng các thư viện khác như MobX hoặc Vuex.
- Sử dụng `React Router` để quản lý routing trong ứng dụng, không sử dụng các thư viện khác như Vue Router hoặc Angular Router.
- Sử dụng `CSS Modules` hoặc `styled-components` để tạo kiểu cho các component, không sử dụng các thư viện CSS khác như Bootstrap hoặc Material UI. 
- Sử dụng `npm` hoặc `yarn` để quản lý các package, không sử dụng các trình quản lý package khác như Bower hoặc Composer.
